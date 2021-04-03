import setuptools

with open("README.md", "r", encoding="utf8") as fh:
    long_description = fh.read()

setuptools.setup(
    name="spreadsheet-tools",
    version="0.0.1",
    author="Fatih Mete",
    author_email="fatihmete@live.com",
    description="All-in-one spreadsheet tools.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/fatihmete/spreadsheet-tools",
    include_package_data=True,
    packages = ['st','st.bin','st.widgets'],
    entry_points = {
        "console_scripts": [
            "spreadsheet-tools = st.bin.__main__:main",
        ]
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
   install_requires=[
          'pandas >= 1.1.0', 
          'openpyxl >= 3.0.5', 
          'PyQt5-sip >= 12.8.1', 
          'PyQt5 >= 5.15.1',
          "numpy >= 1.19.3"
   ],
   python_requires='>=3.6',
)