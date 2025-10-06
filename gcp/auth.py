"""
Google Cloud Platform 認証モジュール

Workload Identity を使用したGoogle Cloud認証を提供します。
"""

from typing import Tuple, Optional
from google.auth import default
from google.auth.credentials import Credentials


def get_gcp_credentials(scopes: Optional[list] = None) -> Tuple[Credentials, str]:
    """
    Google Cloud Platform の認証情報を取得
    
    Args:
        scopes: 必要なスコープのリスト。Noneの場合はデフォルトスコープを使用
    
    Returns:
        tuple: (credentials, project_id)
        
    Raises:
        google.auth.exceptions.DefaultCredentialsError: 認証情報が見つからない場合
    """
    if scopes is None:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.readonly", 
            "https://www.googleapis.com/auth/drive.file",
        ]
    
    credentials, project = default(scopes=scopes)
    return credentials, project


def validate_credentials(credentials: Credentials) -> bool:
    """
    認証情報の有効性を確認
    
    Args:
        credentials: 確認する認証情報
        
    Returns:
        bool: 有効な場合True
    """
    try:
        if not credentials.valid:
            credentials.refresh(None)
        return credentials.valid
    except Exception:
        return False