from setuptools import setup
from os import path

here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'README.md')) as f:
    long_description = f.read()

setup(
    name='pptx-pandas',

    version='0.4',
    
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
                      #prettypandas - for formatted tables
                      "prettypandas @ https://github.com/hottwaj/PrettyPandas/archive/0.0.4jc.tar.gz", 
                      #plotly, for plotly_pandas - wrapped plotly js charts
                      "plotly @ https://github.com/hottwaj/plotly.py/archive/v4.14.0a-jc.tar.gz#egg=plotly&subdirectory=packages/python/plotly",
                      "python-pptx>=0.6.18",
                      "pandas>=1.2.0"],
)

