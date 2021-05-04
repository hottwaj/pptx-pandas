from setuptools import setup
from os import path

here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'README.md')) as f:
    long_description = f.read()

setup(
    name='pptx-pandas',

    version='0.14',
    
    python_requires='>3.8',

    description='Helper methods to convert pandas DataFrames, plotly and matplotlib charts to pptx equivalents',
    long_description=long_description,

    url='https://github.com/hottwaj/pptx-pandas',

    author='Jonathan Clarke',
    author_email='jonathan.a.clarke@gmail.com',

    license='MIT, Copyright 2021',

    classifiers=[
    ],

    keywords='',

    py_modules=['pptx_pandas'],

    extras_require={
        "PrettyPandas": ["PrettyPandas @ https://github.com/hottwaj/PrettyPandas/archive/master.zip"], # for formatted tables
        "plotly-pandas": ["plotly-pandas @ https://github.com/hottwaj/plotly-pandas/archive/main.zip"], # for wrapped plotly charts
    },
    
    install_requires=["python-pptx>=0.6.18",
                      "pandas>=1.2.0"],
)

