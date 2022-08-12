import concurrent.futures
import re
import time
from typing import Generator
from threading import Event

import pywinauto
import win32com.client as win32
from comtypes.client import CreateObject, GetActiveObject

from .. import settings
from .doc import Doc


class App:
    """A class representing a running instance of a QlikView app."""

    def __init__(self) -> None:
        """Initialize an instance."""
        try:  # if QlikView is open try comtypes.client.GetActiveObject() to connect
            self.com: object = GetActiveObject("QlikTech.QlikView")
        # Qlik is open, but comtypes.lsclient.GetActiveObject() failed - use win32 to connect
        except AttributeError:
            self.com: object = win32.DispatchEx("QlikTech.QlikView")
        except OSError:  # Qlik is closed
            try:  # try comtypes.client.CreateObject() to open the app
                self.com: object = CreateObject("QlikTech.QlikView")
            except AttributeError:  # comtypes.client.CreateObject() failed - use win32 to open
                self.com: object = win32.DispatchEx("QlikTech.QlikView")

        # specify keyword directly (write path=) for connect to work
        self.uia: object = pywinauto.Application(backend="uia").connect(
            path=settings.QLIKVIEW_PATH
        )
        self.pid: int = self.com.GetProcessId()
        self.servers = (settings.BMK_SERVER_URL, settings.GIPPO_SERVER_URL)

    def list_docs(self) -> list[str]:
        """Get an iterable of all docs."""
        docs: list[str] = []
        for server in self.servers:
            server_doc_list = self.com.GetServerDocList(server)  # IArrayOfDocListEntry
            for i in range(server_doc_list.Count):
                doc_name = server_doc_list.Item(i).DocName
                docs.append(f"{server}/{doc_name}")
        return docs

    def find_docs(self, pattern: str) -> list[str]:
        """Find all Docs matching the pattern."""
        return [doc for doc in self.list_docs() if re.search(pattern, doc)]

    def _call_open_doc(self, doc_login_path: str, event: Event) -> object:
        """Call an OpenDoc method from the app's COM object."""
        doc =  self.com.OpenDoc(doc_login_path, settings.QLIKVIEW_USER, settings.QLIKVIEW_USER_PASSWORD)
        event.set()
        return doc

    def _login_to_doc(self, event: Event) -> None:
        """Perform a sequence of GUI actions for logging in to the document.

        NOTE: server login is not necessary if we have already preformed a connection
        # to the server during this session.
        """
        # call until the event is set in case 
        # the window did not catch keys upon first call
        while not event.is_set():
            # any errors raised here will nut be catched since 
            # we are not accessing the future's result in ThreadPoolExecutor
            if self.uia.top_window()[f"Password"].exists():
                self.uia.top_window()[f"Password"].Edit.type_keys(
                    settings.QLIKVIEW_SERVER_PASSWORD
                )
                self.uia.top_window()["OK"].click()

            self.uia.top_window()["User Identification"].type_keys(settings.QLIKVIEW_USER)
            self.uia.top_window()["OK"].click()
            self.uia.top_window()["Password"].type_keys(settings.QLIKVIEW_USER_PASSWORD)
            self.uia.top_window()["OK"].click()

    def open_doc(self, doc_path: str) -> object:
        """Open a QlikView document and return a Document object, then create a QlikDoc object based on it."""

        doc_login_path = f"{settings.QLIKVIEW_PROTOCOL}{settings.QLIKVIEW_SERVER_LOGIN}@{doc_path.replace(settings.QLIKVIEW_PROTOCOL, '')}"

        # we have to use multi threading in order to be able to use GUI automation while calling OpenDoc()
        doc_opened = Event()
        with concurrent.futures.ThreadPoolExecutor() as ex:
            future_doc = ex.submit(
                self._call_open_doc,
                doc_login_path,
                doc_opened
            )
            ex.submit(self._login_to_doc, doc_opened)

        return Doc(doc_path, future_doc.result())
