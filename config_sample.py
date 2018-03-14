
"""
Sample config file.
Copy this file to config.py and change the settings below.
"""

__all__ = ['mysql_config']

mysql_config = {
  'user': 'mysql_user',
  'password': 'mysql_password',
  'host': 'localhost',
  'database': 'myqsl_database_name',
  'raise_on_warnings': True,
  'use_pure': True,
  #'use_pure': False,  # Setting use_pure to False causes the connection
                       # to use the C Extension if your Connector/Python
                       # installation includes it. See:
  # https://dev.mysql.com/doc/connector-python/en/connector-python-cext.html
}

