from setuptools import setup, find_packages

with open('README.md') as f:
    readme = f.read()

with open('LICENSE') as f:
    license = f.read()

setup(name='python-dxsr',
		version='0.1',
		description='Package used for performing search/replace-operations on MS Word 2007+ documents',
		long_description=readme,
		url='https://github.com/satvug/python-docx-search-replace',
		author='Gustav Jensen',
		# author_email='n/a',
		license=license,
		# packages=['dxsr'],
		install_requires=[ 'lxml' ],
		packages=find_packages(exclude=('tests', 'docs'))
)
