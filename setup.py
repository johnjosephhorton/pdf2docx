from setuptools import setup, find_packages

setup(
    name="pfd2docx",
    version="0.1",
    packages=find_packages(),
    install_requires=[
        "pdf2image",
        "Pillow",
        "python-docx",
    ],
    entry_points={
        "console_scripts": [
            "pdf2docx=pdf2docx:main",
        ],
    },
)
