# Salesforce Database Sync

This project was created to retrieve Salesforce data points from a PostgreSQL database. The data is then output to an excel file in a format to see trends overtime. The datacontains opportunity counts and totals attached to individual partners at specific dates. 


## Requirements
- PostgreSQL Database
- Salesforce credentials

```
pip install psycopg2-binary simple-salesforce openpyxl
```

config.json JSON configuration file 

```
{
    "database":{
        "user": "user",
        "password": "pass",
        "database": "name"
    },
    "salesforce":{
        "username": "user",
        "password": "password",
        "token": "token"
    }
}
```
