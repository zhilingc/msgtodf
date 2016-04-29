from setuptools import setup
git = "https://github.com/zhilingc/msgtodf"
setup(
    name=msgtodf,
    version=0.1,
    description="Extracts emails and attachments saved in Microsoft Outlook's .msg files to pandas dataframe",
    long_description= "Modified from Matthew Walker's ExtractMsg.",
    url=git,
    download_url="%s/archives/master" % git,
    author='Chen Zhiling',
    author_email='chnzhlng@gmail.com',
    license='GPL',
    scripts=['msgtodf.py'],
    py_modules=['msgtodf'],
    install_requires=['olefile','pandas' ],
)