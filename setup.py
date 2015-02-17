from setuptools import setup, find_packages

setup(
    name = "django-admin-export",
    version = "2.0",
    author = "David Burke",
    author_email = "david@burkesoftware.com",
    description = ("Generic export action for Django admin interface"),
    license = "BSD",
    keywords = "django admin",
    url = "https://github.com/burke-software/django-admin-export",
    packages=find_packages(),
    include_package_data=True,
    classifiers=[
        "Development Status :: 4 - Beta",
        'Environment :: Web Environment',
        'Framework :: Django',
        'Programming Language :: Python',
        'Intended Audience :: Developers',
        'Intended Audience :: System Administrators',
        "License :: OSI Approved :: BSD License",
    ],
    install_requires=['django-report-utils'],
)
