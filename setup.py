from setuptools import setup
from os import path

here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'README.md')) as f:
    long_description = f.read()

setup(
    name='pptx-pandas',

    version='0.1',
    
    python_requires='>3.7',

    description='Helper methods to convert pandas DataFrames, plotly and matplotlib charts to pptx equivalents',
    long_description=long_description,

    url='https://github.com/hottwaj/classproperties',

    author='Jonathan Clarke',
    author_email='jonathan.a.clarke@gmail.com',

    license='MIT, Copyright 2021',

    classifiers=[
    ],

    keywords='',

    #packages=["pptx_pandas"],
    py_modules=['pptx_pandas'],
    
    install_requires=["six",
                      "python-pptx>=0.6.18",
                      "pandas>=1.2.0"],
)
