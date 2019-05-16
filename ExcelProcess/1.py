class Tk(Misc, Wm):
    """Toplevel widget of Tk which represents mostly the main window
    of an application. It has an associated Tcl interpreter."""
    _w = '.'
    def __init__(self, screenName=None, baseName=None, className='Tk',
                 useTk=1, sync=0, use=None):
        """Return a new Toplevel widget on screen SCREENNAME. A new Tcl interpreter will
        be created. BASENAME will be used for the identification of the profile file (see
        readprofile).
        It is constructed from sys.argv[0] without extensions if None is given. CLASSNAME
        is the name of the widget class."""
        self.master = None
        self.children = {}
        self._tkloaded = 0
        # to avoid recursions in the getattr code in case of failure, we
        # ensure that self.tk is always _something_.
        self.tk = None
        if baseName is None:
            import os
            baseName = os.path.basename(sys.argv[0])
            baseName, ext = os.path.splitext(baseName)
            if ext not in ('.py', '.pyc'):
                baseName = baseName + ext
        interactive = 0
        self.tk = _tkinter.create(screenName, baseName, className, interactive, wantobjects, useTk, sync, use)
        if useTk:
            self._loadtk()
        if not sys.flags.ignore_environment:
            # Issue #16248: Honor the -E flag to avoid code injection.
            self.readprofile(baseName, className)
    def loadtk(self):
        if not self._tkloaded:
            self.tk.loadtk()
            self._loadtk()
    def _loadtk(self):
        self._tkloaded = 1
        global _default_root
        # Version sanity checks
        tk_version = self.tk.getvar('tk_version')
        if tk_version != _tkinter.TK_VERSION:
            raise RuntimeError("tk.h version (%s) doesn't match libtk.a version (%s)"
                               % (_tkinter.TK_VERSION, tk_version))
        # Under unknown circumstances, tcl_version gets coerced to float
        tcl_version = str(self.tk.getvar('tcl_version'))
        if tcl_version != _tkinter.TCL_VERSION:
            raise RuntimeError("tcl.h version (%s) doesn't match libtcl.a version (%s)" \
                               % (_tkinter.TCL_VERSION, tcl_version))
        # Create and register the tkerror and exit commands
        # We need to inline parts of _register here, _ register
        # would register differently-named commands.
        if self._tclCommands is None:
            self._tclCommands = []
        self.tk.createcommand('tkerror', _tkerror)
        self.tk.createcommand('exit', _exit)
        self._tclCommands.append('tkerror')
        self._tclCommands.append('exit')
        if _support_default_root and not _default_root:
            _default_root = self
        self.protocol("WM_DELETE_WINDOW", self.destroy)
    def destroy(self):
        """Destroy this and all descendants widgets. This will
        end the application of this Tcl interpreter."""
        for c in list(self.children.values()): c.destroy()
        self.tk.call('destroy', self._w)
        Misc.destroy(self)
        global _default_root
        if _support_default_root and _default_root is self:
            _default_root = None
    def readprofile(self, baseName, className):
        """Internal function. It reads BASENAME.tcl and CLASSNAME.tcl into
        the Tcl Interpreter and calls exec on the contents of BASENAME.py and
        CLASSNAME.py if such a file exists in the home directory."""
        import os
        if 'HOME' in os.environ: home = os.environ['HOME']
        else: home = os.curdir
        class_tcl = os.path.join(home, '.%s.tcl' % className)
        class_py = os.path.join(home, '.%s.py' % className)
        base_tcl = os.path.join(home, '.%s.tcl' % baseName)
        base_py = os.path.join(home, '.%s.py' % baseName)
        dir = {'self': self}
        exec('from tkinter import *', dir)
        if os.path.isfile(class_tcl):
            self.tk.call('source', class_tcl)
        if os.path.isfile(class_py):
            exec(open(class_py).read(), dir)
        if os.path.isfile(base_tcl):
            self.tk.call('source', base_tcl)
        if os.path.isfile(base_py):
            exec(open(base_py).read(), dir)
    def report_callback_exception(self, exc, val, tb):
        """Report callback exception on sys.stderr.

        Applications may want to override this internal function, and
        should when sys.stderr is None."""
        import traceback
        print("Exception in Tkinter callback", file=sys.stderr)
        sys.last_type = exc
        sys.last_value = val
        sys.last_traceback = tb
        traceback.print_exception(exc, val, tb)
    def __getattr__(self, attr):
        "Delegate attribute access to the interpreter object"
        return getattr(self.tk, attr)

# Ideally, the classes Pack, Place and Grid disappear, the
# pack/place/grid methods are defined on the Widget class, and
# everybody uses w.pack_whatever(...) instead of Pack.whatever(w,
# ...), with pack(), place() and grid() being short for
# pack_configure(), place_configure() and grid_columnconfigure(), and
# forget() being short for pack_forget().  As a practical matter, I'm
# afraid that there is too much code out there that may be using the
# Pack, Place or Grid class, so I leave them intact -- but only as
# backwards compatibility features.  Also note that those methods that
# take a master as argument (e.g. pack_propagate) have been moved to
# the Misc class (which now incorporates all methods common between
# toplevel and interior widgets).  Again, for compatibility, these are
# copied into the Pack, Place or Grid class.


def Tcl(screenName=None, baseName=None, className='Tk', useTk=0):
    return Tk(screenName, baseName, className, useTk)



class Pack:
    """Geometry manager Pack.

    Base class to use the methods pack_* in every widget."""
    def pack_configure(self, cnf={}, **kw):
        """Pack a widget in the parent widget. Use as options:
        after=widget - pack it after you have packed widget
        anchor=NSEW (or subset) - position widget according to
                                  given direction
        before=widget - pack it before you will pack widget
        expand=bool - expand widget if parent size grows
        fill=NONE or X or Y or BOTH - fill widget if widget grows
        in=master - use master to contain this widget
        in_=master - see 'in' option description
        ipadx=amount - add internal padding in x direction
        ipady=amount - add internal padding in y direction
        padx=amount - add padding in x direction
        pady=amount - add padding in y direction
        side=TOP or BOTTOM or LEFT or RIGHT -  where to add this widget.
        """
        self.tk.call(
              ('pack', 'configure', self._w)
              + self._options(cnf, kw))
    pack = configure = config = pack_configure
    def pack_forget(self):
        """Unmap this widget and do not use it for the packing order."""
        self.tk.call('pack', 'forget', self._w)
    forget = pack_forget
    def pack_info(self):
        """Return information about the packing options
        for this widget."""
        d = _splitdict(self.tk, self.tk.call('pack', 'info', self._w))
        if 'in' in d:
            d['in'] = self.nametowidget(d['in'])
        return d
    info = pack_info
    propagate = pack_propagate = Misc.pack_propagate
    slaves = pack_slaves = Misc.pack_slaves

class Place:
    """Geometry manager Place.

    Base class to use the methods place_* in every widget."""
    def place_configure(self, cnf={}, **kw):
        """Place a widget in the parent widget. Use as options:
        in=master - master relative to which the widget is placed
        in_=master - see 'in' option description
        x=amount - locate anchor of this widget at position x of master
        y=amount - locate anchor of this widget at position y of master
        relx=amount - locate anchor of this widget between 0.0 and 1.0
                      relative to width of master (1.0 is right edge)
        rely=amount - locate anchor of this widget between 0.0 and 1.0
                      relative to height of master (1.0 is bottom edge)
        anchor=NSEW (or subset) - position anchor according to given direction
        width=amount - width of this widget in pixel
        height=amount - height of this widget in pixel
        relwidth=amount - width of this widget between 0.0 and 1.0
                          relative to width of master (1.0 is the same width
                          as the master)
        relheight=amount - height of this widget between 0.0 and 1.0
                           relative to height of master (1.0 is the same
                           height as the master)
        bordermode="inside" or "outside" - whether to take border width of
                                           master widget into account
        """
        self.tk.call(
              ('place', 'configure', self._w)
              + self._options(cnf, kw))
    place = configure = config = place_configure
    def place_forget(self):
        """Unmap this widget."""
        self.tk.call('place', 'forget', self._w)
    forget = place_forget
    def place_info(self):
        """Return information about the placing options
        for this widget."""
        d = _splitdict(self.tk, self.tk.call('place', 'info', self._w))
        if 'in' in d:
            d['in'] = self.nametowidget(d['in'])
        return d
    info = place_info
    slaves = place_slaves = Misc.place_slaves

class Grid:
    """Geometry manager Grid.

    Base class to use the methods grid_* in every widget."""
    # Thanks to Masazumi Yoshikawa (yosikawa@isi.edu)
    def grid_configure(self, cnf={}, **kw):
        """Position a widget in the parent widget in a grid. Use as options:
        column=number - use cell identified with given column (starting with 0)
        columnspan=number - this widget will span several columns
        in=master - use master to contain this widget
        in_=master - see 'in' option description
        ipadx=amount - add internal padding in x direction
        ipady=amount - add internal padding in y direction
        padx=amount - add padding in x direction
        pady=amount - add padding in y direction
        row=number - use cell identified with given row (starting with 0)
        rowspan=number - this widget will span several rows
        sticky=NSEW - if cell is larger on which sides will this
                      widget stick to the cell boundary
        """
        self.tk.call(
              ('grid', 'configure', self._w)
              + self._options(cnf, kw))
    grid = configure = config = grid_configure
    bbox = grid_bbox = Misc.grid_bbox
    columnconfigure = grid_columnconfigure = Misc.grid_columnconfigure
    def grid_forget(self):
        """Unmap this widget."""
        self.tk.call('grid', 'forget', self._w)
    forget = grid_forget
    def grid_remove(self):
        """Unmap this widget but remember the grid options."""
        self.tk.call('grid', 'remove', self._w)
    def grid_info(self):
        """Return information about the options
        for positioning this widget in a grid."""
        d = _splitdict(self.tk, self.tk.call('grid', 'info', self._w))
        if 'in' in d:
            d['in'] = self.nametowidget(d['in'])
        return d
    info = grid_info
    location = grid_location = Misc.grid_location
    propagate = grid_propagate = Misc.grid_propagate
    rowconfigure = grid_rowconfigure = Misc.grid_rowconfigure
    size = grid_size = Misc.grid_size
    slaves = grid_slaves = Misc.grid_slaves

class BaseWidget(Misc):
    """Internal class."""
    def _setup(self, master, cnf):
        """Internal function. Sets up information about children."""
        if _support_default_root:
            global _default_root
            if not master:
                if not _default_root:
                    _default_root = Tk()
                master = _default_root
        self.master = master
        self.tk = master.tk
        name = None
        if 'name' in cnf:
            name = cnf['name']
            del cnf['name']
        if not name:
            name = self.__class__.__name__.lower()
            if master._last_child_ids is None:
                master._last_child_ids = {}
            count = master._last_child_ids.get(name, 0) + 1
            master._last_child_ids[name] = count
            if count == 1:
                name = '!%s' % (name,)
            else:
                name = '!%s%d' % (name, count)
        self._name = name
        if master._w=='.':
            self._w = '.' + name
        else:
            self._w = master._w + '.' + name
        self.children = {}
        if self._name in self.master.children:
            self.master.children[self._name].destroy()
        self.master.children[self._name] = self
    def __init__(self, master, widgetName, cnf={}, kw={}, extra=()):
        """Construct a widget with the parent widget MASTER, a name WIDGETNAME
        and appropriate options."""
        if kw:
            cnf = _cnfmerge((cnf, kw))
        self.widgetName = widgetName
        BaseWidget._setup(self, master, cnf)
        if self._tclCommands is None:
            self._tclCommands = []
        classes = [(k, v) for k, v in cnf.items() if isinstance(k, type)]
        for k, v in classes:
            del cnf[k]
        self.tk.call(
            (widgetName, self._w) + extra + self._options(cnf))
        for k, v in classes:
            k.configure(self, v)
    def destroy(self):
        """Destroy this and all descendants widgets."""
        for c in list(self.children.values()): c.destroy()
        self.tk.call('destroy', self._w)
        if self._name in self.master.children:
            del self.master.children[self._name]
        Misc.destroy(self)
    def _do(self, name, args=()):
        # XXX Obsolete -- better use self.tk.call directly!
        return self.tk.call((self._w, name) + args)

class Widget(BaseWidget, Pack, Place, Grid):
    """Internal class.

    Base class for a widget which can be positioned with the geometry managers
    Pack, Place or Grid."""
    pass

class Toplevel(BaseWidget, Wm):
    """Toplevel widget, e.g. for dialogs."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct a toplevel widget with the parent MASTER.

        Valid resource names: background, bd, bg, borderwidth, class,
        colormap, container, cursor, height, highlightbackground,
        highlightcolor, highlightthickness, menu, relief, screen, takefocus,
        use, visual, width."""
        if kw:
            cnf = _cnfmerge((cnf, kw))
        extra = ()
        for wmkey in ['screen', 'class_', 'class', 'visual',
                  'colormap']:
            if wmkey in cnf:
                val = cnf[wmkey]
                # TBD: a hack needed because some keys
                # are not valid as keyword arguments
                if wmkey[-1] == '_': opt = '-'+wmkey[:-1]
                else: opt = '-'+wmkey
                extra = extra + (opt, val)
                del cnf[wmkey]
        BaseWidget.__init__(self, master, 'toplevel', cnf, {}, extra)
        root = self._root()
        self.iconname(root.iconname())
        self.title(root.title())
        self.protocol("WM_DELETE_WINDOW", self.destroy)

class Button(Widget):
    """Button widget."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct a button widget with the parent MASTER.

        STANDARD OPTIONS

            activebackground, activeforeground, anchor,
            background, bitmap, borderwidth, cursor,
            disabledforeground, font, foreground
            highlightbackground, highlightcolor,
            highlightthickness, image, justify,
            padx, pady, relief, repeatdelay,
            repeatinterval, takefocus, text,
            textvariable, underline, wraplength

        WIDGET-SPECIFIC OPTIONS

            command, compound, default, height,
            overrelief, state, width
        """
        Widget.__init__(self, master, 'button', cnf, kw)

    def flash(self):
        """Flash the button.

        This is accomplished by redisplaying
        the button several times, alternating between active and
        normal colors. At the end of the flash the button is left
        in the same normal/active state as when the command was
        invoked. This command is ignored if the button's state is
        disabled.
        """
        self.tk.call(self._w, 'flash')

    def invoke(self):
        """Invoke the command associated with the button.

        The return value is the return value from the command,
        or an empty string if there is no command associated with
        the button. This command is ignored if the button's state
        is disabled.
        """
        return self.tk.call(self._w, 'invoke')

class Canvas(Widget, XView, YView):
    """Canvas widget to display graphical elements like lines or text."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct a canvas widget with the parent MASTER.

        Valid resource names: background, bd, bg, borderwidth, closeenough,
        confine, cursor, height, highlightbackground, highlightcolor,
        highlightthickness, insertbackground, insertborderwidth,
        insertofftime, insertontime, insertwidth, offset, relief,
        scrollregion, selectbackground, selectborderwidth, selectforeground,
        state, takefocus, width, xscrollcommand, xscrollincrement,
        yscrollcommand, yscrollincrement."""
        Widget.__init__(self, master, 'canvas', cnf, kw)
    def addtag(self, *args):
        """Internal function."""
        self.tk.call((self._w, 'addtag') + args)
    def addtag_above(self, newtag, tagOrId):
        """Add tag NEWTAG to all items above TAGORID."""
        self.addtag(newtag, 'above', tagOrId)
    def addtag_all(self, newtag):
        """Add tag NEWTAG to all items."""
        self.addtag(newtag, 'all')
    def addtag_below(self, newtag, tagOrId):
        """Add tag NEWTAG to all items below TAGORID."""
        self.addtag(newtag, 'below', tagOrId)
    def addtag_closest(self, newtag, x, y, halo=None, start=None):
        """Add tag NEWTAG to item which is closest to pixel at X, Y.
        If several match take the top-most.
        All items closer than HALO are considered overlapping (all are
        closests). If START is specified the next below this tag is taken."""
        self.addtag(newtag, 'closest', x, y, halo, start)
    def addtag_enclosed(self, newtag, x1, y1, x2, y2):
        """Add tag NEWTAG to all items in the rectangle defined
        by X1,Y1,X2,Y2."""
        self.addtag(newtag, 'enclosed', x1, y1, x2, y2)
    def addtag_overlapping(self, newtag, x1, y1, x2, y2):
        """Add tag NEWTAG to all items which overlap the rectangle
        defined by X1,Y1,X2,Y2."""
        self.addtag(newtag, 'overlapping', x1, y1, x2, y2)
    def addtag_withtag(self, newtag, tagOrId):
        """Add tag NEWTAG to all items with TAGORID."""
        self.addtag(newtag, 'withtag', tagOrId)
    def bbox(self, *args):
        """Return a tuple of X1,Y1,X2,Y2 coordinates for a rectangle
        which encloses all items with tags specified as arguments."""
        return self._getints(
            self.tk.call((self._w, 'bbox') + args)) or None
    def tag_unbind(self, tagOrId, sequence, funcid=None):
        """Unbind for all items with TAGORID for event SEQUENCE  the
        function identified with FUNCID."""
        self.tk.call(self._w, 'bind', tagOrId, sequence, '')
        if funcid:
            self.deletecommand(funcid)
    def tag_bind(self, tagOrId, sequence=None, func=None, add=None):
        """Bind to all items with TAGORID at event SEQUENCE a call to function FUNC.

        An additional boolean parameter ADD specifies whether FUNC will be
        called additionally to the other bound function or whether it will
        replace the previous function. See bind for the return value."""
        return self._bind((self._w, 'bind', tagOrId),
                  sequence, func, add)
    def canvasx(self, screenx, gridspacing=None):
        """Return the canvas x coordinate of pixel position SCREENX rounded
        to nearest multiple of GRIDSPACING units."""
        return self.tk.getdouble(self.tk.call(
            self._w, 'canvasx', screenx, gridspacing))
    def canvasy(self, screeny, gridspacing=None):
        """Return the canvas y coordinate of pixel position SCREENY rounded
        to nearest multiple of GRIDSPACING units."""
        return self.tk.getdouble(self.tk.call(
            self._w, 'canvasy', screeny, gridspacing))
    def coords(self, *args):
        """Return a list of coordinates for the item given in ARGS."""
        # XXX Should use _flatten on args
        return [self.tk.getdouble(x) for x in
                           self.tk.splitlist(
                   self.tk.call((self._w, 'coords') + args))]
    def _create(self, itemType, args, kw): # Args: (val, val, ..., cnf={})
        """Internal function."""
        args = _flatten(args)
        cnf = args[-1]
        if isinstance(cnf, (dict, tuple)):
            args = args[:-1]
        else:
            cnf = {}
        return self.tk.getint(self.tk.call(
            self._w, 'create', itemType,
            *(args + self._options(cnf, kw))))
    def create_arc(self, *args, **kw):
        """Create arc shaped region with coordinates x1,y1,x2,y2."""
        return self._create('arc', args, kw)
    def create_bitmap(self, *args, **kw):
        """Create bitmap with coordinates x1,y1."""
        return self._create('bitmap', args, kw)
    def create_image(self, *args, **kw):
        """Create image item with coordinates x1,y1."""
        return self._create('image', args, kw)
    def create_line(self, *args, **kw):
        """Create line with coordinates x1,y1,...,xn,yn."""
        return self._create('line', args, kw)
    def create_oval(self, *args, **kw):
        """Create oval with coordinates x1,y1,x2,y2."""
        return self._create('oval', args, kw)
    def create_polygon(self, *args, **kw):
        """Create polygon with coordinates x1,y1,...,xn,yn."""
        return self._create('polygon', args, kw)
    def create_rectangle(self, *args, **kw):
        """Create rectangle with coordinates x1,y1,x2,y2."""
        return self._create('rectangle', args, kw)
    def create_text(self, *args, **kw):
        """Create text with coordinates x1,y1."""
        return self._create('text', args, kw)
    def create_window(self, *args, **kw):
        """Create window with coordinates x1,y1,x2,y2."""
        return self._create('window', args, kw)
    def dchars(self, *args):
        """Delete characters of text items identified by tag or id in ARGS (possibly
        several times) from FIRST to LAST character (including)."""
        self.tk.call((self._w, 'dchars') + args)
    def delete(self, *args):
        """Delete items identified by all tag or ids contained in ARGS."""
        self.tk.call((self._w, 'delete') + args)
    def dtag(self, *args):
        """Delete tag or id given as last arguments in ARGS from items
        identified by first argument in ARGS."""
        self.tk.call((self._w, 'dtag') + args)
    def find(self, *args):
        """Internal function."""
        return self._getints(
            self.tk.call((self._w, 'find') + args)) or ()
    def find_above(self, tagOrId):
        """Return items above TAGORID."""
        return self.find('above', tagOrId)
    def find_all(self):
        """Return all items."""
        return self.find('all')
    def find_below(self, tagOrId):
        """Return all items below TAGORID."""
        return self.find('below', tagOrId)
    def find_closest(self, x, y, halo=None, start=None):
        """Return item which is closest to pixel at X, Y.
        If several match take the top-most.
        All items closer than HALO are considered overlapping (all are
        closest). If START is specified the next below this tag is taken."""
        return self.find('closest', x, y, halo, start)
    def find_enclosed(self, x1, y1, x2, y2):
        """Return all items in rectangle defined
        by X1,Y1,X2,Y2."""
        return self.find('enclosed', x1, y1, x2, y2)
    def find_overlapping(self, x1, y1, x2, y2):
        """Return all items which overlap the rectangle
        defined by X1,Y1,X2,Y2."""
        return self.find('overlapping', x1, y1, x2, y2)
    def find_withtag(self, tagOrId):
        """Return all items with TAGORID."""
        return self.find('withtag', tagOrId)
    def focus(self, *args):
        """Set focus to the first item specified in ARGS."""
        return self.tk.call((self._w, 'focus') + args)
    def gettags(self, *args):
        """Return tags associated with the first item specified in ARGS."""
        return self.tk.splitlist(
            self.tk.call((self._w, 'gettags') + args))
    def icursor(self, *args):
        """Set cursor at position POS in the item identified by TAGORID.
        In ARGS TAGORID must be first."""
        self.tk.call((self._w, 'icursor') + args)
    def index(self, *args):
        """Return position of cursor as integer in item specified in ARGS."""
        return self.tk.getint(self.tk.call((self._w, 'index') + args))
    def insert(self, *args):
        """Insert TEXT in item TAGORID at position POS. ARGS must
        be TAGORID POS TEXT."""
        self.tk.call((self._w, 'insert') + args)
    def itemcget(self, tagOrId, option):
        """Return the resource value for an OPTION for item TAGORID."""
        return self.tk.call(
            (self._w, 'itemcget') + (tagOrId, '-'+option))
    def itemconfigure(self, tagOrId, cnf=None, **kw):
        """Configure resources of an item TAGORID.

        The values for resources are specified as keyword
        arguments. To get an overview about
        the allowed keyword arguments call the method without arguments.
        """
        return self._configure(('itemconfigure', tagOrId), cnf, kw)
    itemconfig = itemconfigure
    # lower, tkraise/lift hide Misc.lower, Misc.tkraise/lift,
    # so the preferred name for them is tag_lower, tag_raise
    # (similar to tag_bind, and similar to the Text widget);
    # unfortunately can't delete the old ones yet (maybe in 1.6)
    def tag_lower(self, *args):
        """Lower an item TAGORID given in ARGS
        (optional below another item)."""
        self.tk.call((self._w, 'lower') + args)
    lower = tag_lower
    def move(self, *args):
        """Move an item TAGORID given in ARGS."""
        self.tk.call((self._w, 'move') + args)
    def postscript(self, cnf={}, **kw):
        """Print the contents of the canvas to a postscript
        file. Valid options: colormap, colormode, file, fontmap,
        height, pageanchor, pageheight, pagewidth, pagex, pagey,
        rotate, width, x, y."""
        return self.tk.call((self._w, 'postscript') +
                    self._options(cnf, kw))
    def tag_raise(self, *args):
        """Raise an item TAGORID given in ARGS
        (optional above another item)."""
        self.tk.call((self._w, 'raise') + args)
    lift = tkraise = tag_raise
    def scale(self, *args):
        """Scale item TAGORID with XORIGIN, YORIGIN, XSCALE, YSCALE."""
        self.tk.call((self._w, 'scale') + args)
    def scan_mark(self, x, y):
        """Remember the current X, Y coordinates."""
        self.tk.call(self._w, 'scan', 'mark', x, y)
    def scan_dragto(self, x, y, gain=10):
        """Adjust the view of the canvas to GAIN times the
        difference between X and Y and the coordinates given in
        scan_mark."""
        self.tk.call(self._w, 'scan', 'dragto', x, y, gain)
    def select_adjust(self, tagOrId, index):
        """Adjust the end of the selection near the cursor of an item TAGORID to index."""
        self.tk.call(self._w, 'select', 'adjust', tagOrId, index)
    def select_clear(self):
        """Clear the selection if it is in this widget."""
        self.tk.call(self._w, 'select', 'clear')
    def select_from(self, tagOrId, index):
        """Set the fixed end of a selection in item TAGORID to INDEX."""
        self.tk.call(self._w, 'select', 'from', tagOrId, index)
    def select_item(self):
        """Return the item which has the selection."""
        return self.tk.call(self._w, 'select', 'item') or None
    def select_to(self, tagOrId, index):
        """Set the variable end of a selection in item TAGORID to INDEX."""
        self.tk.call(self._w, 'select', 'to', tagOrId, index)
    def type(self, tagOrId):
        """Return the type of the item TAGORID."""
        return self.tk.call(self._w, 'type', tagOrId) or None

class Checkbutton(Widget):
    """Checkbutton widget which is either in on- or off-state."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct a checkbutton widget with the parent MASTER.

        Valid resource names: activebackground, activeforeground, anchor,
        background, bd, bg, bitmap, borderwidth, command, cursor,
        disabledforeground, fg, font, foreground, height,
        highlightbackground, highlightcolor, highlightthickness, image,
        indicatoron, justify, offvalue, onvalue, padx, pady, relief,
        selectcolor, selectimage, state, takefocus, text, textvariable,
        underline, variable, width, wraplength."""
        Widget.__init__(self, master, 'checkbutton', cnf, kw)
    def deselect(self):
        """Put the button in off-state."""
        self.tk.call(self._w, 'deselect')
    def flash(self):
        """Flash the button."""
        self.tk.call(self._w, 'flash')
    def invoke(self):
        """Toggle the button and invoke a command if given as resource."""
        return self.tk.call(self._w, 'invoke')
    def select(self):
        """Put the button in on-state."""
        self.tk.call(self._w, 'select')
    def toggle(self):
        """Toggle the button."""
        self.tk.call(self._w, 'toggle')

class Entry(Widget, XView):
    """Entry widget which allows displaying simple text."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct an entry widget with the parent MASTER.

        Valid resource names: background, bd, bg, borderwidth, cursor,
        exportselection, fg, font, foreground, highlightbackground,
        highlightcolor, highlightthickness, insertbackground,
        insertborderwidth, insertofftime, insertontime, insertwidth,
        invalidcommand, invcmd, justify, relief, selectbackground,
        selectborderwidth, selectforeground, show, state, takefocus,
        textvariable, validate, validatecommand, vcmd, width,
        xscrollcommand."""
        Widget.__init__(self, master, 'entry', cnf, kw)
    def delete(self, first, last=None):
        """Delete text from FIRST to LAST (not included)."""
        self.tk.call(self._w, 'delete', first, last)
    def get(self):
        """Return the text."""
        return self.tk.call(self._w, 'get')
    def icursor(self, index):
        """Insert cursor at INDEX."""
        self.tk.call(self._w, 'icursor', index)
    def index(self, index):
        """Return position of cursor."""
        return self.tk.getint(self.tk.call(
            self._w, 'index', index))
    def insert(self, index, string):
        """Insert STRING at INDEX."""
        self.tk.call(self._w, 'insert', index, string)
    def scan_mark(self, x):
        """Remember the current X, Y coordinates."""
        self.tk.call(self._w, 'scan', 'mark', x)
    def scan_dragto(self, x):
        """Adjust the view of the canvas to 10 times the
        difference between X and Y and the coordinates given in
        scan_mark."""
        self.tk.call(self._w, 'scan', 'dragto', x)
    def selection_adjust(self, index):
        """Adjust the end of the selection near the cursor to INDEX."""
        self.tk.call(self._w, 'selection', 'adjust', index)
    select_adjust = selection_adjust
    def selection_clear(self):
        """Clear the selection if it is in this widget."""
        self.tk.call(self._w, 'selection', 'clear')
    select_clear = selection_clear
    def selection_from(self, index):
        """Set the fixed end of a selection to INDEX."""
        self.tk.call(self._w, 'selection', 'from', index)
    select_from = selection_from
    def selection_present(self):
        """Return True if there are characters selected in the entry, False
        otherwise."""
        return self.tk.getboolean(
            self.tk.call(self._w, 'selection', 'present'))
    select_present = selection_present
    def selection_range(self, start, end):
        """Set the selection from START to END (not included)."""
        self.tk.call(self._w, 'selection', 'range', start, end)
    select_range = selection_range
    def selection_to(self, index):
        """Set the variable end of a selection to INDEX."""
        self.tk.call(self._w, 'selection', 'to', index)
    select_to = selection_to

class Frame(Widget):
    """Frame widget which may contain other widgets and can have a 3D border."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct a frame widget with the parent MASTER.

        Valid resource names: background, bd, bg, borderwidth, class,
        colormap, container, cursor, height, highlightbackground,
        highlightcolor, highlightthickness, relief, takefocus, visual, width."""
        cnf = _cnfmerge((cnf, kw))
        extra = ()
        if 'class_' in cnf:
            extra = ('-class', cnf['class_'])
            del cnf['class_']
        elif 'class' in cnf:
            extra = ('-class', cnf['class'])
            del cnf['class']
        Widget.__init__(self, master, 'frame', cnf, {}, extra)

class Label(Widget):
    """Label widget which can display text and bitmaps."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct a label widget with the parent MASTER.

        STANDARD OPTIONS

            activebackground, activeforeground, anchor,
            background, bitmap, borderwidth, cursor,
            disabledforeground, font, foreground,
            highlightbackground, highlightcolor,
            highlightthickness, image, justify,
            padx, pady, relief, takefocus, text,
            textvariable, underline, wraplength

        WIDGET-SPECIFIC OPTIONS

            height, state, width

        """
        Widget.__init__(self, master, 'label', cnf, kw)

class Listbox(Widget, XView, YView):
    """Listbox widget which can display a list of strings."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct a listbox widget with the parent MASTER.

        Valid resource names: background, bd, bg, borderwidth, cursor,
        exportselection, fg, font, foreground, height, highlightbackground,
        highlightcolor, highlightthickness, relief, selectbackground,
        selectborderwidth, selectforeground, selectmode, setgrid, takefocus,
        width, xscrollcommand, yscrollcommand, listvariable."""
        Widget.__init__(self, master, 'listbox', cnf, kw)
    def activate(self, index):
        """Activate item identified by INDEX."""
        self.tk.call(self._w, 'activate', index)
    def bbox(self, index):
        """Return a tuple of X1,Y1,X2,Y2 coordinates for a rectangle
        which encloses the item identified by the given index."""
        return self._getints(self.tk.call(self._w, 'bbox', index)) or None
    def curselection(self):
        """Return the indices of currently selected item."""
        return self._getints(self.tk.call(self._w, 'curselection')) or ()
    def delete(self, first, last=None):
        """Delete items from FIRST to LAST (included)."""
        self.tk.call(self._w, 'delete', first, last)
    def get(self, first, last=None):
        """Get list of items from FIRST to LAST (included)."""
        if last is not None:
            return self.tk.splitlist(self.tk.call(
                self._w, 'get', first, last))
        else:
            return self.tk.call(self._w, 'get', first)
    def index(self, index):
        """Return index of item identified with INDEX."""
        i = self.tk.call(self._w, 'index', index)
        if i == 'none': return None
        return self.tk.getint(i)
    def insert(self, index, *elements):
        """Insert ELEMENTS at INDEX."""
        self.tk.call((self._w, 'insert', index) + elements)
    def nearest(self, y):
        """Get index of item which is nearest to y coordinate Y."""
        return self.tk.getint(self.tk.call(
            self._w, 'nearest', y))
    def scan_mark(self, x, y):
        """Remember the current X, Y coordinates."""
        self.tk.call(self._w, 'scan', 'mark', x, y)
    def scan_dragto(self, x, y):
        """Adjust the view of the listbox to 10 times the
        difference between X and Y and the coordinates given in
        scan_mark."""
        self.tk.call(self._w, 'scan', 'dragto', x, y)
    def see(self, index):
        """Scroll such that INDEX is visible."""
        self.tk.call(self._w, 'see', index)
    def selection_anchor(self, index):
        """Set the fixed end oft the selection to INDEX."""
        self.tk.call(self._w, 'selection', 'anchor', index)
    select_anchor = selection_anchor
    def selection_clear(self, first, last=None):
        """Clear the selection from FIRST to LAST (included)."""
        self.tk.call(self._w,
                 'selection', 'clear', first, last)
    select_clear = selection_clear
    def selection_includes(self, index):
        """Return 1 if INDEX is part of the selection."""
        return self.tk.getboolean(self.tk.call(
            self._w, 'selection', 'includes', index))
    select_includes = selection_includes
    def selection_set(self, first, last=None):
        """Set the selection from FIRST to LAST (included) without
        changing the currently selected elements."""
        self.tk.call(self._w, 'selection', 'set', first, last)
    select_set = selection_set
    def size(self):
        """Return the number of elements in the listbox."""
        return self.tk.getint(self.tk.call(self._w, 'size'))
    def itemcget(self, index, option):
        """Return the resource value for an ITEM and an OPTION."""
        return self.tk.call(
            (self._w, 'itemcget') + (index, '-'+option))
    def itemconfigure(self, index, cnf=None, **kw):
        """Configure resources of an ITEM.

        The values for resources are specified as keyword arguments.
        To get an overview about the allowed keyword arguments
        call the method without arguments.
        Valid resource names: background, bg, foreground, fg,
        selectbackground, selectforeground."""
        return self._configure(('itemconfigure', index), cnf, kw)
    itemconfig = itemconfigure

class Menu(Widget):
    """Menu widget which allows displaying menu bars, pull-down menus and pop-up menus."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct menu widget with the parent MASTER.

        Valid resource names: activebackground, activeborderwidth,
        activeforeground, background, bd, bg, borderwidth, cursor,
        disabledforeground, fg, font, foreground, postcommand, relief,
        selectcolor, takefocus, tearoff, tearoffcommand, title, type."""
        Widget.__init__(self, master, 'menu', cnf, kw)
    def tk_popup(self, x, y, entry=""):
        """Post the menu at position X,Y with entry ENTRY."""
        self.tk.call('tk_popup', self._w, x, y, entry)
    def activate(self, index):
        """Activate entry at INDEX."""
        self.tk.call(self._w, 'activate', index)
    def add(self, itemType, cnf={}, **kw):
        """Internal function."""
        self.tk.call((self._w, 'add', itemType) +
                 self._options(cnf, kw))
    def add_cascade(self, cnf={}, **kw):
        """Add hierarchical menu item."""
        self.add('cascade', cnf or kw)
    def add_checkbutton(self, cnf={}, **kw):
        """Add checkbutton menu item."""
        self.add('checkbutton', cnf or kw)
    def add_command(self, cnf={}, **kw):
        """Add command menu item."""
        self.add('command', cnf or kw)
    def add_radiobutton(self, cnf={}, **kw):
        """Addd radio menu item."""
        self.add('radiobutton', cnf or kw)
    def add_separator(self, cnf={}, **kw):
        """Add separator."""
        self.add('separator', cnf or kw)
    def insert(self, index, itemType, cnf={}, **kw):
        """Internal function."""
        self.tk.call((self._w, 'insert', index, itemType) +
                 self._options(cnf, kw))
    def insert_cascade(self, index, cnf={}, **kw):
        """Add hierarchical menu item at INDEX."""
        self.insert(index, 'cascade', cnf or kw)
    def insert_checkbutton(self, index, cnf={}, **kw):
        """Add checkbutton menu item at INDEX."""
        self.insert(index, 'checkbutton', cnf or kw)
    def insert_command(self, index, cnf={}, **kw):
        """Add command menu item at INDEX."""
        self.insert(index, 'command', cnf or kw)
    def insert_radiobutton(self, index, cnf={}, **kw):
        """Addd radio menu item at INDEX."""
        self.insert(index, 'radiobutton', cnf or kw)
    def insert_separator(self, index, cnf={}, **kw):
        """Add separator at INDEX."""
        self.insert(index, 'separator', cnf or kw)
    def delete(self, index1, index2=None):
        """Delete menu items between INDEX1 and INDEX2 (included)."""
        if index2 is None:
            index2 = index1

        num_index1, num_index2 = self.index(index1), self.index(index2)
        if (num_index1 is None) or (num_index2 is None):
            num_index1, num_index2 = 0, -1

        for i in range(num_index1, num_index2 + 1):
            if 'command' in self.entryconfig(i):
                c = str(self.entrycget(i, 'command'))
                if c:
                    self.deletecommand(c)
        self.tk.call(self._w, 'delete', index1, index2)
    def entrycget(self, index, option):
        """Return the resource value of a menu item for OPTION at INDEX."""
        return self.tk.call(self._w, 'entrycget', index, '-' + option)
    def entryconfigure(self, index, cnf=None, **kw):
        """Configure a menu item at INDEX."""
        return self._configure(('entryconfigure', index), cnf, kw)
    entryconfig = entryconfigure
    def index(self, index):
        """Return the index of a menu item identified by INDEX."""
        i = self.tk.call(self._w, 'index', index)
        if i == 'none': return None
        return self.tk.getint(i)
    def invoke(self, index):
        """Invoke a menu item identified by INDEX and execute
        the associated command."""
        return self.tk.call(self._w, 'invoke', index)
    def post(self, x, y):
        """Display a menu at position X,Y."""
        self.tk.call(self._w, 'post', x, y)
    def type(self, index):
        """Return the type of the menu item at INDEX."""
        return self.tk.call(self._w, 'type', index)
    def unpost(self):
        """Unmap a menu."""
        self.tk.call(self._w, 'unpost')
    def xposition(self, index): # new in Tk 8.5
        """Return the x-position of the leftmost pixel of the menu item
        at INDEX."""
        return self.tk.getint(self.tk.call(self._w, 'xposition', index))
    def yposition(self, index):
        """Return the y-position of the topmost pixel of the menu item at INDEX."""
        return self.tk.getint(self.tk.call(
            self._w, 'yposition', index))

class Menubutton(Widget):
    """Menubutton widget, obsolete since Tk8.0."""
    def __init__(self, master=None, cnf={}, **kw):
        Widget.__init__(self, master, 'menubutton', cnf, kw)

class Message(Widget):
    """Message widget to display multiline text. Obsolete since Label does it too."""
    def __init__(self, master=None, cnf={}, **kw):
        Widget.__init__(self, master, 'message', cnf, kw)

class Radiobutton(Widget):
    """Radiobutton widget which shows only one of several buttons in on-state."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct a radiobutton widget with the parent MASTER.

        Valid resource names: activebackground, activeforeground, anchor,
        background, bd, bg, bitmap, borderwidth, command, cursor,
        disabledforeground, fg, font, foreground, height,
        highlightbackground, highlightcolor, highlightthickness, image,
        indicatoron, justify, padx, pady, relief, selectcolor, selectimage,
        state, takefocus, text, textvariable, underline, value, variable,
        width, wraplength."""
        Widget.__init__(self, master, 'radiobutton', cnf, kw)
    def deselect(self):
        """Put the button in off-state."""

        self.tk.call(self._w, 'deselect')
    def flash(self):
        """Flash the button."""
        self.tk.call(self._w, 'flash')
    def invoke(self):
        """Toggle the button and invoke a command if given as resource."""
        return self.tk.call(self._w, 'invoke')
    def select(self):
        """Put the button in on-state."""
        self.tk.call(self._w, 'select')

class Scale(Widget):
    """Scale widget which can display a numerical scale."""
    def __init__(self, master=None, cnf={}, **kw):
        """Construct a scale widget with the parent MASTER.

        Valid resource names: activebackground, background, bigincrement, bd,
        bg, borderwidth, command, cursor, digits, fg, font, foreground, from,
        highlightbackground, highlightcolor, highlightthickness, label,
        length, orient, relief, repeatdelay, repeatinterval, resolution,
        showvalue, sliderlength, sliderrelief, state, takefocus,
        tickinterval, to, troughcolor, variable, width."""
        Widget.__init__(self, master, 'scale', cnf, kw)
    def get(self):
        """Get the current value as integer or float."""
        value = self.tk.call(self._w, 'get')
        try:
            return self.tk.getint(value)
        except (ValueError, TypeError, TclError):
            return self.tk.getdouble(value)
    def set(self, value):
        """Set the value to VALUE."""
        self.tk.call(self._w, 'set', value)
    def coords(self, value=None):
        """Return a tuple (X,Y) of the point along the centerline of the
        trough that corresponds to VALUE or the current value if None is
        given."""

        return self._getints(self.tk.call(self._w, 'coords', value))
    def identify(self, x, y):
        """Return where the point X,Y lies. Valid return values are "slider",
        "though1" and "though2"."""
        return self.tk.call(self._w, 'identify', x, y)
