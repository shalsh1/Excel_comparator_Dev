"""
Versioning module following Semantic Versioning 2.0.0
https://semver.org/

Version format: MAJOR.MINOR.PATCH[-PRERELEASE][+BUILD]
"""

# Semantic Versioning components
__major__ = 0
__minor__ = 0
__patch__ = 2
__prerelease__ = None  # e.g., "alpha", "beta", "rc1" or None for release
__build__ = None  # e.g., "build.123" or None

# Construct full version string
def _build_version():
    """Build version string following Semantic Versioning 2.0.0."""
    version = f"{__major__}.{__minor__}.{__patch__}"
    
    if __prerelease__:
        version += f"-{__prerelease__}"
    
    if __build__:
        version += f"+{__build__}"
    
    return version

# Public version string
__version__ = _build_version()

# Tuple format for easy comparison (MAJOR, MINOR, PATCH)
version_info = (__major__, __minor__, __patch__)

# Convenient access
VERSION = __version__
MAJOR = __major__
MINOR = __minor__
PATCH = __patch__

if __name__ == "__main__":
    print(f"Excel Comparator Tool v{__version__}")
    print(f"Version Info: {version_info}")
