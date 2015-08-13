from setuptools import setup
import os

execfile(os.path.join('sheetsync','version.py'))

with open('README.rst') as fh:
    long_description = fh.read()

with open('requirements.txt') as fh:
    requirements = [line.strip() for line in fh.readlines()]

setup(
    name='sheetsync',
    version=__version__,
    description="Synchronize rows of data with a google spreadsheet",
    long_description=long_description,
    author='Mark Brenig-Jones',
    author_email='markbrenigjones@gmail.com',
    url='https://github.com/mbrenig/SheetSync/',
    packages=['sheetsync'],
    platforms='any',
    install_requires=requirements,
    classifiers=[
        "Development Status :: 4 - Beta",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 2.7",
        ],
)
