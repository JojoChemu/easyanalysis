import subprocess
import sys
import unittest
from unittest.mock import patch

from fastapi import HTTPException
from sqlalchemy.exc import OperationalError

from backend import db
from backend.routers.auth import require_database


class DatabaseAvailabilityTests(unittest.TestCase):
    def test_app_import_does_not_connect_to_database(self):
        script = """
from unittest.mock import patch
from sqlalchemy.engine import Engine
with patch.object(Engine, 'connect', side_effect=AssertionError('Unexpected startup connection')):
    from backend.main import app, health_check
    assert health_check()['status'] == 'ok'
    assert any(route.path == '/api/loan-arrears-difference' for route in app.routes)
"""
        subprocess.run([sys.executable, "-c", script], check=True, capture_output=True)

    def test_failed_initialization_can_recover(self):
        failure = OperationalError(None, None, Exception("database DNS unavailable"))
        with patch.object(db, "_schema_ready", False), patch.object(
            db.Base.metadata, "create_all", side_effect=[failure, None]
        ) as create:
            with self.assertRaises(OperationalError):
                db.initialize_database()
            self.assertFalse(db._schema_ready)
            db.initialize_database()
            db.initialize_database()
            self.assertEqual(create.call_count, 2)

    def test_account_outage_returns_safe_retryable_error(self):
        failure = OperationalError(None, None, Exception("private connection details"))
        with patch("backend.routers.auth.initialize_database", side_effect=failure):
            with self.assertRaises(HTTPException) as error:
                next(require_database())
        self.assertEqual(error.exception.status_code, 503)
        self.assertNotIn("private connection", error.exception.detail)
        self.assertEqual(error.exception.headers["Retry-After"], "10")

    def test_query_outage_is_also_handled(self):
        with patch("backend.routers.auth.initialize_database"):
            dependency = require_database()
            next(dependency)
            with self.assertRaises(HTTPException) as error:
                dependency.throw(OperationalError(None, None, Exception("connection lost")))
            self.assertEqual(error.exception.status_code, 503)


if __name__ == "__main__":
    unittest.main()
