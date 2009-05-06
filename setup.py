from setuptools import setup, find_packages
import sys, os

version = '0.1'

setup(name='ooo2tools.core',
      version=version,
      description="Tools to handle openoffice.org in listen mode",
      long_description=open("README.txt").read() + "\n" +
                       open(os.path.join("docs", "HISTORY.txt")).read(),
      classifiers=[],  # Get strings from http://pypi.python.org/pypi?%3Aaction=list_classifiers
      keywords='open office pyuno',
      author='Jean-Michel FRANCOIS',
      author_email='jeanmichel.francois@makina-corpus.org',
      url='http://pypi.python.org/pypi/ooo2tools.core',
      license='GPL',
      packages=find_packages(exclude=['ez_setup',]),
      packages=find_packages('src'),
      package_dir = {'': 'src'},
      namespace_packages=['ooo2tools'],
      zip_safe=False,
      install_requires=[
          # -*- Extra requirements: -*-
      ],
      entry_points="""
      # -*- Entry points: -*-
      """,
      )

# vim:set et sts=4 ts=4 tw=80:
