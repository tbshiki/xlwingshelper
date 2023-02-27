from setuptools import setup, find_packages
import xlwingshelper

VERSION = xlwingshelper.__version__

setup(name="xlwingshelper", version=VERSION, packages=find_packages())
