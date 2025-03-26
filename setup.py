from setuptools import setup, find_packages

setup(
    name="mseqauto",
    version="0.1.0",
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'ind-sort=mseqauto.scripts.ind_sort_files:main',
            'ind-mseq=mseqauto.scripts.ind_auto_mseq:main',
            'ind-zip=mseqauto.scripts.ind_zip_files:main',
            'plate-sort=mseqauto.scripts.plate_sort_files:main',
            'plate-mseq=mseqauto.scripts.plate_auto_mseq:main',
            'plate-zip=mseqauto.scripts.plate_zip_files:main',
            'validate-zip=mseqauto.scripts.validate_zip_files:main',
        ],
    },
    install_requires=[
        'pywinauto>=0.6.8',
        'numpy>=1.20.0',
        'openpyxl>=3.0.7',
        'pylightxl>=1.60',
        'pywin32>=300',
        'psutil>=5.9.0',
    ],
    extras_require={
        'dev': [
            'pytest>=6.0.0',
            'pytest-cov>=2.12.0',
            'black>=21.5b2',
            'flake8>=3.9.2',
        ],
    },
    python_requires='>=3.6',
    author="Tyler Rudig",
    author_email="tyler@functionalbio.com",
    description="Automation tools for DNA sequencing workflow",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/Rhomboidal1/MseqAuto",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Topic :: Scientific/Engineering :: Bio-Informatics",
    ],
)