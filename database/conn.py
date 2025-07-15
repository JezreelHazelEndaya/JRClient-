from sqlalchemy import create_engine, func, distinct, cast, text, desc
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import scoping, sessionmaker
from dotenv import load_dotenv
import os
load_dotenv()

class DBManager(object):
    def __init__(self):
        self.connection = self.create_connection_string()
        options = {
            "pool_recycle": 3600,
            "pool_size": 10,
            "pool_timeout": 30,
            "max_overflow": 30,
            "echo": False,
            "execution_options": {"autocommit": True},
        }

        self.engine = create_engine(self.connection, **options)
        self.DBSession = scoping.scoped_session(sessionmaker(bind=self.engine, ))

    @property
    def session(self):
        return self.DBSession()

    @staticmethod
    def create_connection_string() -> str:
        username = os.environ.get("DB_USERNAME")
        password = os.environ.get("DB_PASSWORD")
        host = os.environ.get("SQL_HOST")
        port = os.environ.get("SQL_PORT")
        db_name = os.environ.get("DB_NAME")
        return f"postgresql://{username}:{password}@{host}:{port}/{db_name}"



