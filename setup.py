import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name='excel_tools',
    version='1.0',
    author="Jean-Paul Mieville",
    author_email="jpmieville@gmail.com",
    description="A library to simplify the work with excel sheets",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/jpmieville/excel_tools.git",
    packages=setuptools.find_packages(),
    install_requires=["xlrd",
                      "xlwt",
                      "dateutil",
                      "win32com"
                      ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
