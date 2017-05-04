from setuptools import setup

setup(name='sheetparser',
      version='0.1',
      description='Turns an Excel (or other..) workbook made of sheets containing several tables into a usable source of data',
      url='http://github.com/gcoffin/sheetparser',
      author='Guillaume Coffin',
      author_email='guill.coffin@gmail.com',
      license='LGPLv3',
      packages=['sheetparser','sheetparser.backends','sheetparser.tests'],
      package_data = {
        'sheetparser.tests': ['test_table1.xlsx']
        },
      install_requires = ['six'],
      classifiers=[
        # How mature is this project? Common values are
        #   3 - Alpha
        #   4 - Beta
        #   5 - Production/Stable
        'Development Status :: 3 - Alpha',
        
        # Indicate who your project is intended for
        'Intended Audience :: Developers',
        
        # Pick your license as you wish (should match "license" above)
        'License :: OSI Approved :: GNU Lesser General Public License v3 (LGPLv3)',
        
        # Specify the Python versions you support here. In particular, ensure
        # that you indicate whether you support Python 2, Python 3 or both.
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        
        'Operating System :: MacOS :: MacOS X',
        'Operating System :: Microsoft :: Windows',
        'Operating System :: POSIX',

        ],
      keywords='excel tables parsing'
)
