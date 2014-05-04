from setuptools import setup

with open('README.rst') as fh:
    long_description = fh.read()

setup(
    name='SheetSync',
    version='0.1',
    description="Synchronize rows of data with a google spreadsheet",
    long_description=long_description,
    author='Mark Brenig-Jones',
    author_email='markbrenigjones@gmail.com',
    url='https://github.com/mbrenig/SheetSync/',
    packages=['sheetsync'],
    platforms='any',
    install_requires=['gdata>=2.0.18','python-dateutil>=1.5'],
    classifiers=[
        "Development Status :: 1 - Planning",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 2.7",
        "Programming Language :: Python :: 3",
        ],
)
