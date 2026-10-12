"""
CAD-Cowork-MCP: MCP server for controlling CAD software via COM interface.
"""

__version__ = "2.0.0"

import os
import json
import logging

logger = logging.getLogger('cad_mcp')


def load_config():
    """Tải tập tin cấu hình"""
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        logger.info("Tệp cấu hình được tải thành công")
        return config
    except Exception as e:
        logger.error(f"Không tải được tập tin cấu hình: {str(e)}")
        return {
            "server": {
                "name": "CAD MCP Server",
                "version": "2.0.0",
            },
            "cad": {
                "type": "AUTO",
                "startup_wait_time": 20,
                "command_delay": 0.5
            },
            "output": {
                "directory": "output",
                "default_filename": "cad_drawing.dwg"
            }
        }


config = load_config()

__all__ = ['config', '__version__']
