import os
from dotenv import load_dotenv
import mysql.connector
from mysql.connector import pooling
from datetime import datetime
import logging
from typing import Optional
from contextlib import contextmanager

# Load environment variables from .env file
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class Database:
    def __init__(self):
        # Load from environment variables
        self.host = os.getenv("DB_HOST")
        self.user = os.getenv("DB_USER")
        self.password = os.getenv("DB_PASSWORD")
        self.database = os.getenv("DB_NAME")
        self.pool_size = int(os.getenv("DB_POOL_SIZE", "10"))
        
        self.pool = pooling.MySQLConnectionPool(
            pool_name="assignment_pool",
            pool_size=self.pool_size,
            pool_reset_session=True,
            host=self.host,
            user=self.user,
            password=self.password,
            database=self.database
        )

    @contextmanager
    def get_connection(self):
        conn = self.pool.get_connection()
        try:
            yield conn
        finally:
            conn.close()

    def insert_assignment(self, assignment_type: str) -> int:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            try:
                query = """
                    INSERT INTO assignment 
                    (assignment_date, assignment_type) 
                    VALUES (%s, %s)
                """
                current_date = datetime.now().strftime('%Y-%m-%d')
                cursor.execute(query, (current_date, assignment_type))
                conn.commit()
                assignment_id = cursor.lastrowid
                logger.info(f"Successfully inserted assignment record with Assignment ID: {assignment_id}")
                return assignment_id
            except mysql.connector.Error as err:
                logger.error(f"Error inserting assignment record: {err}")
                conn.rollback()
                raise
            finally:
                cursor.close()

    def insert_assignment_detail(self, assignment_id: int, word_count: int, due_date: str, 
                               assignment_type: str, software_required: str, topic: str, 
                               university_name: str, citation_style: str = "") -> int:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            try:
                query = """
                    INSERT INTO assignment_detail 
                    (assignment_id, word_count, due_date, assignment_type, software_required, topic, university_name, citation_style) 
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                """
                cursor.execute(query, (assignment_id, word_count, due_date, assignment_type, 
                                     software_required, topic, university_name, citation_style))
                conn.commit()
                detail_id = cursor.lastrowid
                logger.info(f"Successfully inserted assignment detail with Assignment ID: {assignment_id}")
                return detail_id
            except mysql.connector.Error as err:
                logger.error(f"Error inserting assignment detail: {err}")
                conn.rollback()
                raise
            finally:
                cursor.close()

    def insert_assignment_instruction(self, assignment_id: int, instruction: str, rubric: str, static_instruction: str, additional_information: str = "") -> int:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            try:
                query = """
                    INSERT INTO assignment_instruction
                    (assignment_id, instruction, rubric, static_instruction, additional_information)
                    VALUES (%s, %s, %s, %s, %s)
                """
                cursor.execute(query, (assignment_id, instruction, rubric, static_instruction, additional_information))
                conn.commit()
                instruction_id = cursor.lastrowid
                logger.info(f"Successfully inserted assignment instruction with Assignment ID: {assignment_id}")
                return instruction_id
            except mysql.connector.Error as err:
                logger.error(f"Error inserting assignment instruction: {err}")
                conn.rollback()
                raise
            finally:
                cursor.close()

    def insert_assignment_material(self, assignment_id: int, actual_filename: str, document_path: str, helping_material: str) -> int:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            try:
                query = """
                    INSERT INTO assignment_material 
                    (assignment_id, actual_filename, document_path, helping_material) 
                    VALUES (%s, %s, %s, %s)
                """
                cursor.execute(query, (assignment_id, actual_filename, document_path, helping_material))
                conn.commit()
                material_id = cursor.lastrowid
                logger.info(f"Successfully inserted assignment material with Assignment ID: {assignment_id}")
                return material_id
            except mysql.connector.Error as err:
                logger.error(f"Error inserting assignment material: {err}")
                conn.rollback()
                raise
            finally:
                cursor.close()

    def insert_assignment_text(self, assignment_id: int, filename: str, brief: str, draft: str) -> int:
        with self.get_connection() as conn:
            cursor = conn.cursor()
            try:
                query = """
                    INSERT INTO assignment_text 
                    (assignment_id, filename, brief, draft) 
                    VALUES (%s, %s, %s, %s)
                """
                cursor.execute(query, (assignment_id, filename, brief, draft))
                conn.commit()
                text_id = cursor.lastrowid
                logger.info(f"Successfully inserted assignment text with Assignment ID: {assignment_id}")
                return text_id
            except mysql.connector.Error as err:
                logger.error(f"Error inserting assignment text: {err}")
                conn.rollback()
                raise
            finally:
                cursor.close()

    def get_assignment_by_id(self, assignment_id):
        with self.get_connection() as conn:
            cursor = conn.cursor(dictionary=True)
            try:
                # assignment table
                cursor.execute("SELECT * FROM assignment WHERE assignment_id = %s", (assignment_id,))
                assignment = cursor.fetchone()
                if not assignment:
                    return None

                # assignment_text table
                cursor.execute("SELECT brief, draft FROM assignment_text WHERE assignment_id = %s", (assignment_id,))
                text = cursor.fetchone()
                assignment["brief"] = text["brief"] if text and text["brief"] else ""
                assignment["outline"] = text["draft"] if text and text["draft"] else ""

                # assignment_material table
                cursor.execute("SELECT helping_material FROM assignment_material WHERE assignment_id = %s", (assignment_id,))
                material = cursor.fetchone()
                assignment["helping_material"] = material["helping_material"] if material and material["helping_material"] else ""

                return assignment
            except Exception as e:
                logger.error(f"Error in get_assignment_by_id: {e}")
                return None
            finally:
                cursor.close()

    def close(self):
        # No-op for pool, but kept for compatibility
        pass