from setuptools import setup, find_packages

setup(
    name="StatPrint",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "python-docx>=0.8.11",
        "fpdf>=1.7.2",
        "pandas>=1.0.0",
    ],
    author="Nolan Leavitt",
    author_email="leavittnolan20@gmail.com",
    description="A report generation library for Word and PDF documents",
    url="https://github.com/npleavitt/statprint",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)
