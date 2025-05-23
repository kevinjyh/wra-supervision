import os
import warnings
from pathlib import Path
from typing import Dict, List, Literal, Optional

import xlwings as xw
from pydantic import UUID4, computed_field
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """See .env.template for documentation"""

    def __init__(self, **values):
        super().__init__(**values)
        if self.public_addin_store is not None:
            warnings.warn(
                "The 'XLWINGS_PUBLIC_ADDIN_STORE' field is deprecated and will be removed in "
                "future versions. Use 'XLWINGS_CDN_OFFICEJS' instead.",
                DeprecationWarning,
            )
            self.cdn_officejs = self.public_addin_store

    model_config = SettingsConfigDict(
        env_prefix="XLWINGS_",
        env_file=os.getenv(
            "DOTENV_PATH", Path(__file__).parent.parent.resolve() / ".env"
        ),
        extra="ignore",
    )
    add_security_headers: bool = True
    auth_providers: Optional[List[str]] = []
    auth_required_roles: Optional[List[str]] = []
    auth_entraid_client_id: Optional[str] = None
    auth_entraid_tenant_id: Optional[str] = None
    auth_entraid_multitenant: bool = False
    app_path: str = ""
    base_dir: Path = Path(__file__).resolve().parent
    object_cache_url: Optional[str] = None
    object_cache_expire_at: Optional[str] = "0 12 * * sat"
    object_cache_enable_compression: bool = True
    cors_allow_origins: List[str] = []
    date_format: Optional[str] = None
    enable_alpinejs_csp: bool = True
    enable_bootstrap: bool = True
    enable_examples: bool = True
    enable_excel_online: bool = True
    enable_hotreload: bool = True
    enable_htmx: bool = True
    enable_socketio: bool = True
    enable_tests: bool = False
    enable_lite: bool = False
    environment: Literal["dev", "qa", "uat", "staging", "prod"] = "prod"
    functions_namespace: str = "XLWINGS"
    hostname: Optional[str] = None
    is_official_lite_addin: Optional[bool] = False
    cdn_pyodide: bool = True
    cdn_officejs: bool = False
    log_level: str = "INFO"
    # These UUIDs will be overwritten by: python run.py init
    manifest_id_dev: UUID4 = "b4b053bc-4798-4b41-821c-9ffe720514ac"
    manifest_id_qa: UUID4 = "0d0be8b1-8d53-4c68-b32e-331f4883bcc2"
    manifest_id_uat: UUID4 = "7ac007bf-c3b6-41a4-9ea5-54f3d387a591"
    manifest_id_staging: UUID4 = "207f35b0-a5b5-4f3f-926d-e82fed7974c0"
    manifest_id_prod: UUID4 = "f764b1b5-5352-4012-8bfd-71dbfa9fa489"
    project_name: str = "xlwings Server"
    public_addin_store: Optional[bool] = None  # Deprecated. Use cdn_officejs instead.
    request_timeout: Optional[int] = 300  # in seconds
    secret_key: Optional[str] = None
    socketio_message_queue_url: Optional[str] = None
    socketio_server_app: bool = False
    static_url_path: str = "/static"
    license_key: Optional[str] = ""
    xlwings_version: str = xw.__version__

    @computed_field
    @property
    def static_dir(self) -> Path:
        return self.base_dir / "static"

    @computed_field
    @property
    def jsconfig(self) -> Dict:
        return {
            "authProviders": self.auth_providers,
            "appPath": self.app_path,
            "xlwingsVersion": self.xlwings_version,
            "onLite": self.enable_lite,
            "isOfficialLiteAddin": self.is_official_lite_addin,
            "requestTimeout": self.request_timeout,
        }


settings = Settings()

# TODO: refactor once xlwings offers a runtime config
if settings.license_key and not os.getenv("XLWINGS_LICENSE_KEY"):
    os.environ["XLWINGS_LICENSE_KEY"] = settings.license_key

if settings.date_format:
    os.environ["XLWINGS_DATE_FORMAT"] = settings.date_format
