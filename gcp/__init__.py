"""
GCP (Google Cloud Platform) integration module

This module provides interfaces for Google Cloud services including:
- Google Sheets API
- Google Drive API  
- Authentication using Workload Identity
"""

from .auth import get_gcp_credentials
from .sheets_client import SheetsClient

__all__ = ['get_gcp_credentials', 'SheetsClient']