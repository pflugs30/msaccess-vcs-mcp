"""
Access COM connection management.

Adapted from db-inspector-mcp with ownership tracking and cleanup patterns.
Manages COM connections to Access databases with proper resource cleanup.

Key Complexity Areas:
- Multiple connection strategies based on whether Access is already running
- Ownership tracking to avoid interfering with user's Access instances
- COM cleanup challenges and garbage collection issues
"""

import os
import re
import sys
from typing import Any

try:
    import win32com.client
    from win32com.client import gencache
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False


def _paths_match(a: str, b: str) -> bool:
    """Case-insensitive, normalised path comparison."""
    try:
        return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))
    except (OSError, ValueError):
        return False


class AccessConnection:
    """
    Manages COM connection to Access database.
    
    Uses ownership tracking to determine cleanup responsibility:
    - If connecting to user's existing Access instance, we don't close it
    - If we create our own Access instance, we're responsible for cleanup
    """
    
    def __init__(self, db_path: str):
        """
        Initialize Access COM connection manager.
        
        Args:
            db_path: Path to Access database file (.accdb, .accda, .mdb)
        """
        if not COM_AVAILABLE:
            raise ImportError(
                "pywin32 is required for COM automation. "
                "Install it with: pip install pywin32"
            )
        
        self._db_path = db_path
        self._app = None
        self._db = None
        
        # Ownership tracking flags - determine cleanup responsibility
        self._owns_app = False  # True only if we created Access via Dispatch
        self._owns_db = False   # True only if we opened db via DBEngine
        self._db_opened_via_getobject = False  # True if connected to existing instance
    
    def _get_access_app(self):
        """
        Get or create Access COM application.
        
        Uses a multi-step approach:
        1. Try GetObject(path) to connect to our database if already open
        2. Try EnsureDispatch -- but check whether it returned an existing
           user instance (has a different database open) vs a fresh one
        3. If EnsureDispatch attached to a user instance with a different
           DB, fall back to DispatchEx to create a guaranteed-isolated
           Access process
        
        IMPORTANT: Access can only have ONE database open at a time.
        - If user has OUR database open -> connect to their instance
        - If user has DIFFERENT database open -> create our OWN isolated instance
        - If no Access running -> create new instance
        
        The ownership check after EnsureDispatch prevents the dangerous
        scenario where close() calls Quit() on the user's Access window
        (ported from db-inspector-mcp's DispatchEx fallback pattern).
        """
        if self._app is None:
            try:
                self._app = win32com.client.GetObject(self._db_path)
                self._db_opened_via_getobject = True
                self._owns_app = False
            except Exception:
                self._app = self._create_or_reuse_instance()
                self._db_opened_via_getobject = False
        return self._app

    def _create_or_reuse_instance(self):
        """Create or attach to an Access instance with correct ownership.

        EnsureDispatch("Access.Application") can silently return an
        existing user-owned instance instead of creating a new one.
        We must check whether the returned instance already has a
        database open to set _owns_app correctly.
        """
        app = gencache.EnsureDispatch("Access.Application")

        try:
            existing_db = app.CurrentDb()
        except Exception:
            existing_db = None

        if existing_db is not None:
            if _paths_match(existing_db.Name, self._db_path):
                self._owns_app = False
                return app
            db_name = os.path.basename(self._db_path)
            other_name = os.path.basename(existing_db.Name)
            print(
                f"[{db_name}] EnsureDispatch returned instance with "
                f"'{other_name}' open -- launching isolated instance via DispatchEx",
                file=sys.stderr,
            )
            return self._create_isolated_instance()

        self._owns_app = True
        return app

    def _create_isolated_instance(self):
        """Create a guaranteed-isolated Access process via DispatchEx.

        Unlike EnsureDispatch, DispatchEx always spawns a new COM server
        process independent of any existing Access instance.
        """
        app = win32com.client.DispatchEx("Access.Application")
        app.UserControl = True
        self._owns_app = True
        return app
    
    def _get_current_db(self):
        """
        Get database object for DAO operations.
        
        Strategy:
        1. If database was opened via GetObject (Access already had it open), use CurrentDb()
        2. Otherwise, use DBEngine.OpenDatabase() which is more reliable
        
        IMPORTANT: This method tracks whether we OWN the database connection:
        - If we use CurrentDb() (user's database), we do NOT own it
        - If we open via DBEngine, we OWN it and are responsible for closing
        """
        if self._db is None:
            app = self._get_access_app()
            
            # If Access already had the database open, CurrentDb() should work
            if self._db_opened_via_getobject:
                try:
                    db = app.CurrentDb()
                    if db is not None:
                        self._db = db
                        self._owns_db = False  # User's database - do NOT close it
                        return self._db
                except Exception:
                    pass
            
            # Use DBEngine.OpenDatabase() - more reliable
            try:
                dbe = app.DBEngine
                # Open database in shared mode (Exclusive=False, ReadOnly=False)
                self._db = dbe.OpenDatabase(self._db_path, False, False)
                self._owns_db = True  # We opened this - we're responsible for closing
            except Exception:
                # Try read-only mode
                try:
                    dbe = app.DBEngine
                    self._db = dbe.OpenDatabase(self._db_path, False, True)
                    self._owns_db = True
                except Exception:
                    # Last resort: try direct DAO without Access
                    try:
                        dbe = win32com.client.Dispatch("DAO.DBEngine.120")
                        self._db = dbe.OpenDatabase(self._db_path, False, True)
                        self._owns_db = True
                    except Exception as e:
                        raise RuntimeError(
                            f"Failed to open database '{self._db_path}' via COM. "
                            f"Ensure the database file exists and is not corrupted. Error: {e}"
                        )
        
        return self._db
    
    def connect(self):
        """
        Connect to Access database via COM.
        
        Returns:
            Tuple of (Application, Database) objects
        """
        app = self._get_access_app()
        db = self._get_current_db()
        return app, db
    
    def get_app(self):
        """Get Access Application object."""
        return self._get_access_app()
    
    def get_db(self):
        """Get DAO Database object."""
        return self._get_current_db()
    
    def close(self):
        """
        Close the database connection and cleanup COM resources.
        
        IMPORTANT: This method respects ownership:
        - If we connected to an existing Access instance (via GetObject), we do NOT
          close Access or the database - the user is still using them!
        - If we created our own Access instance, we clean it up properly.
        """
        # Only close the database if we opened it ourselves
        if self._db is not None and self._owns_db:
            try:
                self._db.Close()
            except Exception:
                pass
        self._db = None
        
        # Only quit Access if we created it ourselves
        if self._app is not None and self._owns_app:
            try:
                self._app.CloseCurrentDatabase()
            except Exception:
                pass
            try:
                self._app.Quit()
            except Exception:
                pass
        self._app = None
        
        # Reset ownership flags
        self._owns_app = False
        self._owns_db = False
        self._db_opened_via_getobject = False
    
    def __enter__(self):
        """Context manager entry."""
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        self.close()
        return False
