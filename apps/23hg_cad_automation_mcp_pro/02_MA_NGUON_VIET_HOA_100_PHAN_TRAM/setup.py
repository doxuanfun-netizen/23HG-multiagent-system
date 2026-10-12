import os
from setuptools import setup, find_packages

readme_path = os.path.join(os.path.dirname(__file__), "README.md")
long_description = open(readme_path, encoding="utf-8").read() if os.path.exists(readme_path) else "23HG CAD MCP System"

setup(
    name="cad-cowork-mcp-23hg",
    version="2.0.0",
    author="Nguyễn Bảo Tú (@baotuhg) - 23HG SYSTEM",
    author_email="baotu23hg@gmail.com",
    description="Máy chủ MCP điều khiển và tự động hóa AutoCAD/GstarCAD/ZWCAD 100% tiếng Việt chuyên ngành AEC",
    long_description=long_description,
    long_description_content_type="text/markdown",
    packages=find_packages(),
    include_package_data=True,
    package_data={
        "cad_cowork_mcp": [
            "*.json",
            "*.dll",
            "prompts/*.md"
        ]
    },
    install_requires=[
        "mcp>=1.0.0",
        "pydantic>=2.0.0",
        "pywin32>=228",
        "ezdxf>=1.1.0",
        "openpyxl>=3.1.0"
    ],
    entry_points={
        "console_scripts": [
            "cad-mcp-23hg=cad_cowork_mcp.server:main_cli",
            "cad-cowork-mcp=cad_cowork_mcp.server:main_cli"
        ]
    },
    python_requires=">=3.10",
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Topic :: Scientific/Engineering :: Interface Engine/Protocol Translator",
    ],
)
