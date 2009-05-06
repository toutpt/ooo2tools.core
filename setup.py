from setuptools import setup, find_packages
import sys, os

version = '0.1'

setup(name='ooo2tools.core',
      version=version,
      description="OOo2 tool using pyUno to manage your document",
      long_description="""\
This product allows you to use a server instance of OOo2 and gives you a simple way to build "macro" and to manage your documents. It has been used to build complex PDF documents from many MSWord documents with table of contents, using templates, etc.""",
      classifiers=[],  # Get strings from http://pypi.python.org/pypi?%3Aaction=list_classifiers
      keywords='open office pyuno',
      author='Jean-Michel FRANCOIS',
      author_email='',
      url='http://plone.org/products/ooo2tools',
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
