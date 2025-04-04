import json
import logging
import os
import sqlite3
import sys
from datetime import datetime, timedelta
from logging import exception
from logging.handlers import RotatingFileHandler
from pathlib import Path


LOG_FILENAME = os.path.join(os.getcwd(), 'monitor.log')
stdout_handler = logging.StreamHandler(stream=sys.stdout)
size_handler = RotatingFileHandler(LOG_FILENAME, backupCount=3, encoding='utf-8')
handlers = [stdout_handler, size_handler]
logging.basicConfig(
    level=logging.INFO,
    format='[%(asctime)s] %(levelname)s - %(message)s',
    handlers=handlers,
)
logger = logging.getLogger('mailmonitor')

class Database:

    def __init__(self, connection):
        logger.debug('Connecting to database...')
        self.connection = connection

    def close_connection(self):
        logger.debug('Closing database connection...')
        self.connection.close()

    def update_global_config(self, account_id):
        command = f"update global_config set account_id = ? where active = 1"
        self.connection.execute(command, [account_id])
        self.connection.commit()

    def get_global_config(self):
        columns = [
            "id",
            "verypdf_folder",
            "application_user",
            "application_password",
            "allowed_domains",
            "printer",
            "clean_attachments",
            "email",
            "provider",
            "access_token",
            "refresh_token",
            "expires_at",
            "scope",
            "provider_name",
            "imap_server",
            "imap_port",
            "smtp_server",
            "smtp_port"
        ]
        command = """
                    select gc.id, gc.verypdf_folder, gc.application_user, gc.application_password, gc.allowed_domains, gc.printer,gc.clean_attachments,
                    ac.email, ac.provider,
                    t.access_token, t.refresh_token, t.expires_at, t.scope,
                    ap.name, ap.imap_server, ap.imap_port, ap.smtp_server, ap.smtp_port
                    from global_config gc
                    left join accounts ac on ac.id = gc.account_id
                    left join tokens t on t.account_id = gc.account_id
                    left join oauth_providers ap on ap.id = gc.provider_id
                    """
        global_config_result = self.connection.execute(command).fetchall()
        if global_config_result:
            return {columns[gcr[0]]: gcr[1] for gcr in enumerate(list(global_config_result[0]))}
        else:
            return {}

    def get_microsoft_config(self):
        return self.get_server_config('microsoft')

    def get_google_config(self):
        return self.get_server_config('google')

    def get_server_config(self, server_name):
        try:
            columns_command = f"PRAGMA table_info(oauth_providers);"
            data_command = f"SELECT * FROM oauth_providers where active = 1 and name like ?"
            result_data = self.connection.execute(data_command, [server_name]).fetchall()
            columns_data = self.connection.execute(columns_command).fetchall()
            data = list(result_data[0])
            columns = [colum_data[1] for colum_data in columns_data]
            microsoft_config = {column: dato for column, dato in zip(columns, data)}
            return microsoft_config
        except IndexError as e:
            logger.error(f'No existe la configuración para {server_name}')
            return {}

    def insert_token(self, account_id,data):
        expires_in = data['expires_in']
        expires_at = datetime.now() + timedelta(seconds=expires_in)
        data.update({'account_id':account_id, 'expires_at':expires_at})
        columns = list(data.keys())
        values = list(data.values())
        command = f"INSERT INTO tokens ({','.join(columns)}) values ({(len(values) * '?,')[:-1]})"
        # insert_token_command = f"INSERT INTO tokens (access_token, refresh_token, id_token, token_type, scope, expires_at, account_id) VALUES (?,?,?, ?, ?, ?, ?)"
        self.connection.execute(command,values)
        self.connection.commit()

    def get_oauth_provider_by_name(self, provider_name):
        logger.debug('Getting oauth provider...')
        provider_id_command = f"select id from oauth_providers where name like ?"
        provider_id = self.connection.execute(provider_id_command, [provider_name]).fetchone()
        return provider_id[0]

    def get_oauth_providers(self):
        logger.debug('Getting oauth providers...')
        providers_command = f"select op.id, op.name, op.imap_server, op.imap_port, op.smtp_server, op.smtp_port from oauth_providers op where active = 1"
        providers = self.connection.execute(providers_command).fetchall()
        return providers

    def create_or_get_account(self, user, provider):
        account_id_command = f"SELECT id from accounts where email like ? and active = 1 and provider like ?"
        account_id = self.connection.execute(account_id_command, [user, provider]).fetchall()
        if account_id:
            logger.debug(f'Account {user}:{provider} exist.')
            id = account_id[0][0]
        else:
            logger.debug(f'Account {user}:{provider} does not exist. Creating...')
            provider_id = self.get_oauth_provider_by_name(provider)
            account_create_command = f"INSERT INTO accounts (email, active, provider, provider_id) VALUES (?, ?, ?, ?);"
            account_id = self.connection.execute(account_create_command, [user, True, provider, provider_id])
            self.connection.commit()
            id = account_id.lastrowid
        return id

    def create_or_update_settings(self, settings):
        global_config = self.get_global_config()
        if not global_config:
            settings_columns = ",".join([x[0] for x in settings])
            num_params = "?," * len(settings)
            settings_values = [x[1] for x in settings]
            command = f"insert into global_config ({settings_columns}) values ({num_params[:-1]})"
            self.connection.execute(command, settings_values)
            self.connection.commit()
        else:
            command = "UPDATE global_config SET "
            command_params = " ,".join([f"{x[0]}=?" for x in settings])
            command_were = ' where id = ?;'
            command += command_params + command_were
            settings_values = [x[1] for x in settings]
            settings_values.append(global_config["id"])
            self.connection.execute(command, settings_values)
            self.connection.commit()

    def create_table(self, table_name, columns):
        cols = [' '.join(tups) for tups in columns]
        logger.debug(f"Creating table {table_name}")
        command = f"CREATE TABLE IF NOT EXISTS {table_name} ({','.join(cols)})"
        self.connection.execute(command)
        logger.debug(f"Table {table_name} created")

    def update_table(self, table_name, columns, values, options=None):
        print(f"Updating table {table_name} with {values}")
        if options is None:
            self.connection.execute(f"UPDATE {table_name} SET {','.join(columns)} = {values}")
        else:
            self.connection.execute(f"UPDATE {table_name} SET {','.join(columns)} = {values} {options}")
        print(f"Table {table_name} updated")

    def create_database(self):
        logger.debug('Creating a new database')
        self.create_table("accounts", [
            ("id", "INTEGER", "PRIMARY KEY AUTOINCREMENT"),
            ("email", "TEXT", "NOT NULL"),
            ("provider", "TEXT", "NOT NULL"),
            ("active", "boolean"),
            ("provider_id", "INTEGER", "NOT NULL"),
            ("created_at", "timestamp", "default CURRENT_TIMESTAMP"),
            ("FOREIGN KEY(", "provider_id", ")REFERENCES", "oauth_providers(", "id", ")")
        ])
        self.create_table("tokens", [
            ("id", "INTEGER", "PRIMARY KEY AUTOINCREMENT"),
            ("access_token", "TEXT", "NOT NULL"),
            ("refresh_token", "TEXT", "NOT NULL"),
            ("id_token", "TEXT"),
            ("token_type", "TEXT"),
            ("scope", "TEXT"),
            ("expires_in", "INTEGER", "NOT NULL"),
            ("expires_at", "timestamp", "NOT NULL"),
            ("created_at", "timestamp", "default CURRENT_TIMESTAMP"),
            ("account_id", "INTEGER", "NOT NULL"),
            ("FOREIGN KEY(", "account_id", ")REFERENCES", "accounts(", "id", ")")
        ])
        self.create_table("global_config", [
            ("id", "INTEGER", "PRIMARY KEY AUTOINCREMENT"),
            ("verypdf_folder", "text"),
            ("application_user", "text"),
            ("application_password", "text"),
            ("allowed_domains", "text"),
            ("printer", "text"),
            ("clean_attachments", "boolean"),
            ("account_id", "integer"),
            ("active", "boolean"),
            ("provider_id", "integer"),
            ("created_at", "timestamp", "default CURRENT_TIMESTAMP"),
            ("FOREIGN KEY(", "provider_id", ")REFERENCES", "oauth_providers(", "id", ")")
        ])
        self.create_table('oauth_providers', [
            ("id", "INTEGER", "PRIMARY KEY AUTOINCREMENT"),
            ("name", "text"),
            ("imap_server", "text"),
            ("imap_port", "INTEGER"),
            ("client_id", "text"),
            ("project_id", "text"),
            ("auth_provider_x509_cert_url", "text"),
            ("client_secret", "text"),
            ("scopes", "text"),
            ("auth_uri", "text"),
            ("redirect_uri", "text"),
            ("redirect_uris", "text"),
            ("token_uri", "text"),
            ("user_info_uri", "text"),
            ("tenant_id", "text"),
            ("object_id", "text"),
            ("active", "boolean"),
            ("created_at", "timestamp", "default CURRENT_TIMESTAMP"),
        ])
        self.insert_or_update_oauth_provider()

    def insert_or_update_oauth_provider(self, name=None, server=None, port=None, smtp_server=None, smtp_port=None):
        # For google
        logger.debug('Inserting or updating oauth providers from file')

        try:
            if not self.get_server_config('Otro'):
                logger.debug('Creating other provider')
                command = f"INSERT INTO oauth_providers (name, active) VALUES (?, ?)"
                self.connection.execute(command, ['Otro', True])
                self.connection.commit()
        except Exception as e:
            logger.error('Error creating provider')

        try:
            config_from_json = json.loads(Path("google.json").read_text())
            if self.get_server_config('google'):
                logger.debug('Updating Google config')
            else:
                logger.debug('Creating Google config')
                columns = list(config_from_json.keys())
                values = list(config_from_json.values())
                command = f"INSERT INTO oauth_providers ({','.join(columns)}) values ({(len(values) * '?,')[:-1]})"
                self.connection.execute(command, values)
                self.connection.commit()

        except FileNotFoundError:
            logger.error("No Google config file found")


        # For microsoft

        try:
            config_from_json = json.loads(Path("microsoft.json").read_text())
            if self.get_server_config('microsoft'):
                logger.debug('Updating microsoft config')
            else:
                logger.debug('Creating microsoft config')
                columns = list(config_from_json.keys())
                values = list(config_from_json.values())
                command = f"INSERT INTO oauth_providers ({','.join(columns)}) values ({(len(values) * '?,')[:-1]})"
                self.connection.execute(command, values)
                self.connection.commit()
        except FileNotFoundError:
            logger.error("No Microsoft config file found")

        if name:
            try:
                logger.debug('Updating provider info.')
                command = f"UPDATE oauth_providers SET imap_server = ?, imap_port = ?, smtp_server = ?, smtp_port = ? WHERE name = ?"
                self.connection.execute(command, [server, port, smtp_server, smtp_port, name])
                self.connection.commit()
            except sqlite3.IntegrityError:
                logger.error('Error updating provider data.')