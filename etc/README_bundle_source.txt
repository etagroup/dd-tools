How to bundle the latest source code

1) Place this script in (or point it at) your project directory.
2) Run:

   python bundle_latest_source.py --path /path/to/project --out source_code_bundle.zip

Notes:
- If the project is a Git repo, the zip will include:
  - tracked files (git ls-files)
  - untracked but non-ignored files
  It will exclude the .git directory automatically.
- If the project is not a Git repo, it will zip the directory tree while excluding common build/cache folders.
