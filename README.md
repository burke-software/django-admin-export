django-admin-export
===================

Generic export to XLSX/HTML/CSV action for the Django admin interface.

Meant for fast and simple exports and lets you choose the data to export.

Features
--------
- Drop in application
- Traverse model relations recurssively

django-admin-export is built with [django-report-utils](https://github.com/burke-software/django-report-utils).
For a a full query builder try using [django-report-builder](https://github.com/burke-software/django-report-builder).

Install
-------
1. ``pip install django-admin-export``
2. Add ``admin_export`` to INSTALLED_APPS
3. Add ``url(r'^admin_export/', include("admin_export.urls", namespace="admin_export")),`` to your project's urls.py

Usage
-----
Go to any admin page, select fields, then select the export to xls action. Then check off any fields you want to export.
