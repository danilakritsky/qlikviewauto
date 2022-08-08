"""This module provides an App class, instance of which serves as a point of entry to the QlikView API."""

from .classes.app import App


class App(App):
    def __init__(self):
        super().__init__()
