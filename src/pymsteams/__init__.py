#!/usr/bin/env python

# https://github.com/rveachkc/pymsteams/
# reference: https://dev.outlook.com/connectors/reference
import requests


class TeamsWebhookException(Exception):
    """custom exception for failed webhook call"""

    pass


class cardsection:
    def __init__(self):
        self.payload = {}


class potentialaction:
    def addOpenURI(self, _name, _targets):
        """
        Creates a OpenURI action

        https://docs.microsoft.com/en-us/outlook/actionable-messages/message-card-reference#openuri-action

        :param _name: *Name of the text to appear inside the ActionCard*
        :type _name: str
        :param _targets: *A list of dictionaries, ex: `{"os": "default", "uri": "https://www..."}`*
        :type _targets: list(dict())
        """
        pass

    def __init__(self, _name, _type="ActionCard"):
        self.payload = {}
        self.payload["@type"] = _type
        self.payload["name"] = _name
        self.choices = choice()


class choice:
    def __init__(self):
        self.choices = []


class connectorcard:
    def __init__(
        self, hookurl, http_proxy=None, https_proxy=None, http_timeout=60, verify=None
    ):
        self.payload = {}
        self.hookurl = hookurl
        self.proxies = {}
        self.http_timeout = http_timeout
        self.verify = verify
        self.last_http_response = None

        if http_proxy:
            self.proxies["http"] = http_proxy

        if https_proxy:
            self.proxies["https"] = https_proxy

        if not self.proxies:
            self.proxies = None


class async_connectorcard(connectorcard):
    """Asynchronous connector card for Microsoft Teams."""

    pass
