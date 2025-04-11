#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from setuptools import setup, find_packages
import os

# Read requirements
with open('requirements.txt') as f:
    requirements = f.read().splitlines()

# Read README for long description
with open('README.md', 'r', encoding='utf-8') as f:
    long_description = f.read()

setup(
    name="altium_rule_generator",
    version="1.0.0",
    description="Altium Designer Rule Generator - Excel to RUL Converter",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="AltiumTools",
    author_email="info@altiumtools.com",
    url="https://github.com/altiumtools/altium-rule-generator",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    include_package_data=True,
    install_requires=requirements,
    entry_points={
        'console_scripts': [
            'altium-rule-generator=main:main',
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Topic :: Scientific/Engineering :: Electronic Design Automation (EDA)",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
)
