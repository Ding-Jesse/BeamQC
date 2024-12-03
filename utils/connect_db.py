import os
from pymongo.mongo_client import MongoClient
from pymongo.server_api import ServerApi
from pymongo.database import Database
from flask import current_app, g


def get_db(database_name, uri=None):

    if not uri is None:
        client = MongoClient(uri, server_api=ServerApi('1'))

        db = client.get_database(database_name)
        return db
    if uri is None:
        try:
            uri = current_app.config['MONGO_URL']
        except RuntimeError:
            uri = os.environ['MONGO_URL'].replace('"', '')
            client = MongoClient(uri, server_api=ServerApi('1'))
            db = client.get_database(database_name)
            return db
    if 'db' not in g:
        # Create a new client and connect to the server
        client = MongoClient(uri, server_api=ServerApi('1'))

        db = client.get_database(database_name)

        g.db = db
    return g.db


def add_error_log(db: Database, data: dict, collection_name: str):
    error_log = db.get_collection(collection_name)
    error_log.insert_one(data)


def get_all_error_log(db: Database):
    error_log = db.get_collection('Error Log')
    return list(error_log.find())


def close_db(e=None):
    db = g.pop('db', None)

    if db is not None:
        db.client.close()
