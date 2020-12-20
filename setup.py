from setuptools import setup

setup(
    name='Gama_Inventory',
    version='0.1.0',
    author='Jugal Bilimoria',
    description="An inventory management program for GAMA",
    url='https://github.com/JugalBili/Python-GAMA-Inventory',
    license='license.txt',
    long_description=open('README.md').read(),
    install_requires=[
        "tk >= 8.6.10"
        "gspread = 3.6.0"
        "oauth2client = 4.1.3"
    ],
    python_requires='>=3.6',
)