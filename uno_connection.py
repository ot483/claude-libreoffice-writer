"""UNO connection manager - singleton, lazy connect, auto-reconnect."""

import uno

CONNECT_STR = "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"


class UnoConnection:
    _instance = None

    def __init__(self):
        self._ctx = None
        self._smgr = None
        self._desktop = None

    @classmethod
    def get(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def connect(self):
        """Establish UNO bridge to a running LibreOffice instance."""
        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )
        try:
            self._ctx = resolver.resolve(CONNECT_STR)
        except Exception as e:
            raise RuntimeError(
                "Cannot connect to LibreOffice. "
                "Make sure it's running with: "
                "soffice --accept='socket,host=localhost,port=2002;urp' --norestore"
            ) from e
        self._smgr = self._ctx.ServiceManager
        self._desktop = self._smgr.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self._ctx
        )

    def _ensure_connected(self):
        """Connect if not yet connected, or reconnect if bridge is dead."""
        if self._desktop is None:
            self.connect()
            return
        try:
            # Probe the bridge - this throws if LO was closed
            self._desktop.getCurrentComponent()
        except Exception:
            self._desktop = None
            self.connect()

    @property
    def desktop(self):
        self._ensure_connected()
        return self._desktop

    @property
    def smgr(self):
        self._ensure_connected()
        return self._smgr

    @property
    def ctx(self):
        self._ensure_connected()
        return self._ctx

    def get_document(self):
        """Return the currently active Writer document."""
        doc = self.desktop.getCurrentComponent()
        if doc is None:
            raise RuntimeError("No document is open in LibreOffice.")
        # Verify it's a text document
        if not doc.supportsService("com.sun.star.text.TextDocument"):
            raise RuntimeError(
                "The active document is not a Writer document. "
                "Please open or focus a Writer (.odt/.doc/.docx) file."
            )
        return doc

    def create_instance(self, service_name):
        """Create a UNO service instance."""
        return self._smgr.createInstanceWithContext(service_name, self._ctx)
