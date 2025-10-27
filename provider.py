"""
Document Converter Plugin Provider for Dify
"""

from dify_plugin import Plugin


class DocumentConverterProvider(Plugin):
    def _validate_config(self, credentials: dict) -> None:
        """
        Validate the plugin configuration
        """
        pass