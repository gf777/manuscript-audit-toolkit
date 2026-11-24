from setuptools import setup, find_packages

setup(
    name="manuscript-audit",
    version="2.0.0",
    description=(
        "End-to-end scientific manuscript auditing, AI-prose de-slopping, spreadsheet/figure "
        "forensics, non-destructive OpenXML Track Changes & phrase-anchored comments, and "
        "in-place multi-tab Google Docs sync."
    ),
    author="Giulio Formenti",
    packages=find_packages(exclude=["tests*"]),
    install_requires=[
        "python-docx>=1.1.0",
        "google-genai>=0.1.1",
        "requests>=2.31.0",
        "tqdm>=4.66.0",
        "openpyxl>=3.1.0",
        "Pillow>=10.0.0",
        "numpy>=1.24.0",
    ],
    entry_points={
        "console_scripts": [
            "manuscript-audit=manuscript_audit.cli:main"
        ]
    },
    python_requires=">=3.9",
)
