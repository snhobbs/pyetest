#!/usr/bin/env python3
from setuptools import setup, find_packages

with open("requirements.txt", "r") as f:
    requirements = f.readlines()

with open("README.md", "r") as f:
    description = f.read()

setup(
    name="pyetest",
    version="0.0.1",
    description=description,
    long_description_content_type="text/markdown",
    url="https://github.com/snhobbs/pyetest",
    author="Simon Hobbs",
    author_email="simon.hobbs@electrooptical.net",
    license="MIT",
    packages=find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
    install_requires=requirements,
    test_suite="nose.collector",
    tests_require=["nose"],
    scripts=[],
    include_package_data=True,
    zip_safe=True,
)
