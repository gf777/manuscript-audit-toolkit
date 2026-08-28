from setuptools import setup, find_packages

setup(
    name="manuscript-audit",
    version="1.0.0",
    description="Automated academic manuscript audit, pre/post revision alignment, and native OpenXML comment injection.",
    author="Giulio Formenti",
    packages=find_packages(),
    install_requires=[
        "python-docx>=1.1.0",
        "google-genai>=0.1.1",
        "requests>=2.31.0",
        "tqdm>=4.66.0"
    ],
    entry_points={
        "console_scripts": [
            "manuscript-audit=manuscript_audit.cli:main"
        ]
    },
    python_requires=">=3.9",
)
