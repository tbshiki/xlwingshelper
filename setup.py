from setuptools import setup, find_packages

# 以下の内容を適宜変更してください
NAME = "xlwingshelper"
VERSION = "0.2"
DESCRIPTION = "A helper library for working with Excel using xlwings."
LONG_DESCRIPTION = "A longer description of your project"  # 通常はREADMEファイルから読み込む
LICENSE = "MIT"
CLASSIFIERS = [
    "Programming Language :: Python :: 3",
    "License :: OSI Approved :: MIT License",
    "Operating System :: OS Independent",
]

AUTHOR = "tbshiki"
AUTHOR_EMAIL = "info@tbshiki.com"
URL = "https://github.com/tbshiki/" + NAME
INSTALL_REQUIRES = ["xlwings"]  # 依存するパッケージのリスト

setup(
    name=NAME,
    version=VERSION,
    description=DESCRIPTION,
    long_description=LONG_DESCRIPTION,
    license=LICENSE,
    classifiers=CLASSIFIERS,
    author=AUTHOR,
    author_email=AUTHOR_EMAIL,
    url=URL,
    packages=find_packages(),
    install_requires=INSTALL_REQUIRES,
)
