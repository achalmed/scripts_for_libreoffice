"""
lib/logger.py — Sistema de logging centralizado.

Un único logger de consola para toda la aplicación; --verbose baja el
nivel a DEBUG. Los módulos nunca hacen print() de diagnóstico directo:
así el formato queda consistente y es fácil añadir un file handler.

Author : Edison Achalma (@achalmed)
Version: 2.0.0
"""

import logging


def setup_logger(name: str, verbose: bool = False) -> logging.Logger:
    """
    Configures and returns the application logger.

    The handler is attached only once so repeated calls (e.g. from
    tests) don't duplicate output lines.

    Args:
        name: Logger name (usually the tool name).
        verbose: When True, DEBUG messages are shown.

    Returns:
        The configured logging.Logger instance.
    """
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)

    if not logger.handlers:
        formatter = logging.Formatter(
            "[%(levelname)s] %(asctime)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    return logger
