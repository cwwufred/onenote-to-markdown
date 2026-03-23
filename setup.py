from setuptools import setup, find_packages

setup(
    name="onenote2md",
    version="0.1.0",
    description="Convert OneNote .one files to Markdown",
    author="OneNote2MD",
    packages=find_packages(),
    python_requires=">=3.8",
    install_requires=[
        "click>=8.0.0",
        "customtkinter>=5.2.0",
        "markdown-it-py>=3.0.0",
        "beautifulsoup4>=4.12.0",
        "lxml>=4.9.0",
        "python-dotenv>=1.0.0",
    ],
    entry_points={
        "console_scripts": [
            "onenote2md=onenote2md.cli:main",
        ],
    },
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: End Users/Desktop",
        "Programming Language :: Python :: 3",
    ],
)