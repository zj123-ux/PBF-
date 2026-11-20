"""Web 请求来源校验（纯逻辑，便于单元测试）。

服务只绑 127.0.0.1，但浏览器里的其它站点仍可向本机 POST。仅比较 Origin
hostname 与请求 Host 会被 DNS rebinding 绕过：攻击者让自有域名解析到
127.0.0.1，Origin 与 Host 同为该域名。因此要求请求 Host 与 Origin
都必须显式是本机地址。
"""
from __future__ import annotations

from urllib.parse import urlparse

LOCAL_HOSTS = {"127.0.0.1", "localhost"}


def _hostname(value: str | None) -> str:
    """从 'host:port' / URL 中取小写 hostname；解析失败返回空串。"""
    if not value:
        return ""
    value = value.strip()
    if "://" in value:
        return (urlparse(value).hostname or "").lower()
    return value.rsplit(":", 1)[0].strip("[]").lower()


def origin_allowed(origin: str | None, host_header: str | None,
                   sec_fetch_site: str | None = None) -> bool:
    """是否放行一个 POST 请求。

    - 无 Origin（命令行/脚本调用）放行；显式 cross-site 仍拒绝。
    - 浏览器请求：Host 必须是本机地址，Origin 也必须是本机地址。
    """
    if (sec_fetch_site or "").strip().lower() == "cross-site":
        return False
    if not origin:
        return True
    request_host = _hostname(host_header)
    origin_host = _hostname(origin)
    return request_host in LOCAL_HOSTS and origin_host in LOCAL_HOSTS
