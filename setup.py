from setuptools import setup

setup(
    name='mgtb',
    version='0.1',
    license='BSD 2.0',
    url="https://github.com/starttolearning/mergetables",
    description="A simple script to merger excel files with same template.",
    author='Wilton Lee',
    author_email='wilton_lee@163.com',
    py_modules=['mgtb', ],
    install_requires=[
        'xlrd',
        'xlwt',
    ],
    python_requires=">=3, !=3.0.*, !=3.1.*, !=3.2.*, !=3.3.*",
    entry_points='''
        [console_scripts]
        mgtb=mgtb:mgtb
    ''',
)
