"""Push exact local Git objects through GitHub's API when Git transport is unavailable.

Only fast-forward main updates and new, immutable tags are allowed. Tokens are
read from the environment or Git Credential Manager and are never written out.
"""
from datetime import datetime, timedelta, timezone
import argparse
import base64
import json
import os
import re
import subprocess
import urllib.error
import urllib.request


def git(*args):
    return subprocess.check_output(['git', *args])


def actor(line):
    match = re.fullmatch(r'.+? (.+) <([^>]+)> (\d+) ([+-])(\d{2})(\d{2})', line)
    if not match:
        raise ValueError('Unsupported Git actor metadata')
    minutes = (int(match[5]) * 60 + int(match[6])) * (1 if match[4] == '+' else -1)
    date = datetime.fromtimestamp(int(match[3]), timezone(timedelta(minutes=minutes)))
    return {'name': match[1], 'email': match[2], 'date': date.isoformat()}


class GitHub:
    def __init__(self):
        remote = git('remote', 'get-url', 'origin').decode().strip()
        match = re.fullmatch(r'(?:https://github\.com/|git@github\.com:)([^/]+)/([^/]+?)(?:\.git)?/?', remote)
        if not match:
            raise ValueError('origin must be a github.com repository')
        self.repository = f'{match[1]}/{match[2]}'
        token = os.environ.get('GH_TOKEN') or os.environ.get('GITHUB_TOKEN')
        if not token:
            env = os.environ.copy()
            env.update(GIT_TERMINAL_PROMPT='0', GCM_INTERACTIVE='Never')
            result = subprocess.run(['git', 'credential', 'fill'],
                input=f'url=https://github.com/{self.repository}.git\n\n',
                text=True, capture_output=True, env=env, timeout=30)
            values = dict(line.split('=', 1) for line in result.stdout.splitlines() if '=' in line)
            token = values.get('password') if result.returncode == 0 else None
        if not token:
            raise ValueError('Sign in with GitHub CLI or Git Credential Manager first')
        self.token = token

    def request(self, path, method='GET', data=None):
        req = urllib.request.Request(f'https://api.github.com/repos/{self.repository}/{path}',
            method=method, data=None if data is None else json.dumps(data).encode('utf-8'),
            headers={'Authorization': 'Bearer ' + self.token,
                     'Accept': 'application/vnd.github+json', 'Content-Type': 'application/json',
                     'X-GitHub-Api-Version': '2022-11-28', 'User-Agent': 'QingYanToolbox-Publish'})
        with urllib.request.urlopen(req, timeout=60) as response:
            return json.load(response)

    def head(self):
        return self.request('git/ref/heads/main')['object']['sha']

    def upload_commit(self, sha):
        entries = []
        for row in git('ls-tree', '-rz', sha).split(b'\0'):
            if not row:
                continue
            meta, path = row.split(b'\t', 1)
            mode, kind, blob_sha = meta.decode().split()
            if kind != 'blob':
                raise ValueError('Submodules require the standard Git transport')
            content = git('cat-file', 'blob', blob_sha)
            entry = {'path': path.decode('utf-8'), 'type': 'blob', 'mode': mode}
            try:
                entry['content'] = content.decode('utf-8')
            except UnicodeDecodeError:
                blob = self.request('git/blobs', 'POST', {
                    'content': base64.b64encode(content).decode(), 'encoding': 'base64'})
                if blob['sha'] != blob_sha:
                    raise ValueError('Uploaded binary blob differs from local Git')
                entry['sha'] = blob_sha
            entries.append(entry)
        tree = self.request('git/trees', 'POST', {'tree': entries})
        if tree['sha'] != git('rev-parse', sha + '^{tree}').decode().strip():
            raise ValueError('Uploaded tree differs from local Git')
        raw = git('cat-file', 'commit', sha).decode('utf-8')
        header, message = raw.split('\n\n', 1)
        lines = header.splitlines()
        if any(not line.startswith(('tree ', 'parent ', 'author ', 'committer ')) for line in lines):
            raise ValueError('Signed or extended commits require standard Git transport')
        data = {'tree': tree['sha'], 'message': message,
                'parents': [line[7:] for line in lines if line.startswith('parent ')]}
        for key in ('author', 'committer'):
            data[key] = actor(next(line for line in lines if line.startswith(key + ' ')))
        result = self.request('git/commits', 'POST', data)
        if result['sha'] != sha:
            raise ValueError('Uploaded commit differs from local Git; main was not updated')
        print(f'Uploaded exact commit {sha[:12]}', flush=True)

    def push_tag(self, name, head):
        if not re.fullmatch(r'v\d+\.\d+\.\d+', name):
            raise ValueError('Release tags must be vX.Y.Z')
        if git('rev-parse', name + '^{commit}').decode().strip() != head:
            raise ValueError('Release tag must point at current HEAD')
        tag_sha = git('rev-parse', name).decode().strip()
        try:
            existing = self.request('git/ref/tags/' + name)
        except urllib.error.HTTPError as error:
            if error.code != 404:
                raise
        else:
            if existing['object']['sha'] != tag_sha:
                raise ValueError('Remote tag already exists with a different object')
            return
        if git('cat-file', '-t', name).decode().strip() == 'tag':
            header, message = git('cat-file', 'tag', name).decode('utf-8').split('\n\n', 1)
            lines = header.splitlines()
            if len(lines) != 4 or any(x not in ('object', 'type', 'tag', 'tagger') for x in
                                      [line.split(' ', 1)[0] for line in lines]):
                raise ValueError('Extended tags require the standard Git transport')
            result = self.request('git/tags', 'POST', {'tag': name, 'message': message,
                'object': head, 'type': 'commit',
                'tagger': actor(next(line for line in lines if line.startswith('tagger ')))})
            if result['sha'] != tag_sha:
                raise ValueError('Uploaded tag differs from local Git')
        self.request('git/refs', 'POST', {'ref': 'refs/tags/' + name, 'sha': tag_sha})


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--check', action='store_true', help='Check remote ancestry without publishing')
    parser.add_argument('--tag')
    args = parser.parse_args()
    if git('branch', '--show-current').decode().strip() != 'main':
        raise ValueError('Publish from main')
    api = GitHub()
    remote = api.head()
    head = git('rev-parse', 'HEAD').decode().strip()
    ancestry = subprocess.run(['git', 'merge-base', '--is-ancestor', remote, head], capture_output=True)
    if ancestry.returncode:
        raise ValueError('Fetch and integrate remote changes before publishing; no force update is allowed')
    if args.check:
        print(f'Remote main {remote[:12]} is an ancestor of local HEAD')
        return
    commits = git('rev-list', '--reverse', remote + '..' + head).decode().splitlines()
    for commit in commits:
        api.upload_commit(commit)
    if api.head() != remote:
        raise ValueError('Remote changed during upload; retry after reviewing the changes')
    if head != remote:
        api.request('git/refs/heads/main', 'PATCH', {'sha': head, 'force': False})
    if args.tag:
        api.push_tag(args.tag, head)
    git('update-ref', 'refs/remotes/origin/main', head)
    print(f'Published {head[:12]}' + (f' and {args.tag}' if args.tag else ''))


if __name__ == '__main__':
    try:
        main()
    except urllib.error.HTTPError as error:
        raise SystemExit(f'GitHub API request failed: HTTP {error.code}')
