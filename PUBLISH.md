# Publishing to PyPI

This guide will help you publish `docx2everything` to PyPI.

## Prerequisites

1. **PyPI Account**: Create an account at https://pypi.org/account/register/
2. **TestPyPI Account** (recommended for testing): Create an account at https://test.pypi.org/account/register/
3. **API Tokens**: Generate API tokens for both PyPI and TestPyPI:
   - Go to Account Settings → API tokens
   - Create a new token with appropriate scope (project or account)

## Step 1: Install Build Tools

```bash
pip install --upgrade build twine
```

## Step 2: Clean Previous Builds (if any)

```bash
rm -rf build/ dist/ *.egg-info/
```

## Step 3: Build the Package

```bash
python -m build
```

This will create:
- `dist/docx2everything-1.0.0.tar.gz` (source distribution)
- `dist/docx2everything-1.0.0-py3-none-any.whl` (wheel distribution)

## Step 4: Test on TestPyPI (Recommended)

First, test your package on TestPyPI to ensure everything works:

```bash
# Upload to TestPyPI
twine upload --repository testpypi dist/*

# When prompted, use:
# Username: __token__
# Password: <your-testpypi-api-token>
```

Then test installation from TestPyPI:

```bash
pip install --index-url https://test.pypi.org/simple/ docx2everything
```

## Step 5: Upload to PyPI

Once you've verified everything works on TestPyPI:

```bash
twine upload dist/*

# When prompted, use:
# Username: __token__
# Password: <your-pypi-api-token>
```

Alternatively, you can use environment variables or config files to avoid entering credentials each time.

## Step 6: Verify Installation

After publishing, verify the package is available:

```bash
pip install docx2everything
docx2everything --help
```

## Updating the Package

To publish a new version:

1. Update the version in `pyproject.toml` (and `setup.py` if you want to keep it in sync)
2. Update `__version__` in `docx2everything/__init__.py`
3. Follow steps 2-6 above

## Troubleshooting

- **"File already exists"**: The version number already exists on PyPI. Increment the version.
- **"Invalid distribution"**: Check that all required files are included in MANIFEST.in
- **"Missing metadata"**: Verify pyproject.toml is complete and valid

## Security Notes

- Never commit API tokens to git
- Use API tokens instead of passwords
- Consider using `keyring` for secure credential storage:
  ```bash
  pip install keyring
  twine upload --repository testpypi dist/*  # Will prompt securely
  ```
