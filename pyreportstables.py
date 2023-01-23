"""module for creating PDF reports with text, images, text boxes, and tables.
Relies on and extends matplotlib with dependencies on PIL and PyPDF2 in
Python 3 environments.

Author(s):  Stanton K. Nielson
Date:       January 23, 2023
Version:    1.2

-------------------------------------------------------------------------------
This is free and unencumbered software released into the public domain.

Anyone is free to copy, modify, publish, use, compile, sell, or
distribute this software, either in source code form or as a compiled
binary, for any purpose, commercial or non-commercial, and by any
means.

In jurisdictions that recognize copyright laws, the author or authors
of this software dedicate any and all copyright interest in the
software to the public domain. We make this dedication for the benefit
of the public at large and to the detriment of our heirs and
successors. We intend this dedication to be an overt act of
relinquishment in perpetuity of all present and future rights to this
software under copyright law.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS BE LIABLE FOR ANY CLAIM, DAMAGES OR
OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE,
ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR
OTHER DEALINGS IN THE SOFTWARE.

For more information, please refer to <http://unlicense.org/>
-------------------------------------------------------------------------------
"""


from sys import version, version_info
if version_info.major < 3: 
    raise RuntimeError('Python version {} is not supported'.format(version))

import os, re, math, copy, inspect, tempfile, matplotlib, PIL
from collections.abc import Sequence
from itertools import product
from matplotlib import pyplot as plot
from matplotlib import font_manager
from matplotlib.text import Text as BaseText
from matplotlib.font_manager import FontProperties as BaseFont
from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter


# LOADING OF ALL TRUETYPE SYSTEM FONTS
for font in font_manager.findSystemFonts(fontpaths=None, fontext='ttf'):
    font_manager.fontManager.addfont(font)


# METHOD FOR OVERRIDE
def _get_wrapped_text(self):
    """
    Return a copy of the text with new lines added, so that
    the text is wrapped relative to the parent figure.
    """
    # Used to override matplotlib.text.Text._get_wrapped_text to remove
    # whitespace at end of each wrapped line
    if self.get_usetex():
        return self.get_text()

    # Build the line incrementally, for a more accurate measure of length
    line_width = self._get_wrap_line_width()
    wrapped_lines = []

    # New lines in the user's text force a split
    unwrapped_lines = self.get_text().split('\n')
    unwrapped_lines = [i.rstrip() for i in unwrapped_lines]# GETS TRUE LINES

    # Now wrap each individual unwrapped line
    for unwrapped_line in unwrapped_lines:

        sub_words = unwrapped_line.split(' ')
        # Remove items from sub_words as we go, so stop when empty
        while len(sub_words) > 0:
            if len(sub_words) == 1:
                # Only one word, so just add it to the end
                wrapped_lines.append(sub_words.pop(0))
                continue

            for i in range(2, len(sub_words) + 1):
                # Get width of all words up to and including here
                line = ' '.join(sub_words[:i])
                current_width = self._get_rendered_text_width(line)

                # If all these words are too wide, append all not including
                # last word
                if current_width > line_width:
                    wrapped_lines.append(' '.join(sub_words[:i - 1]))
                    sub_words = sub_words[i - 1:]
                    break

                # Otherwise if all words fit in the width, append them all
                elif i == len(sub_words):
                    wrapped_lines.append(' '.join(sub_words[:i]))
                    sub_words = []
                    break
    
    return '\n'.join(wrapped_lines)


# OVERRIDING METHOD
matplotlib.text.Text._get_wrapped_text = _get_wrapped_text


class BaseCell(matplotlib.table.Cell):

    # CLASS INHERITANCE ALLOWS FOR USE OF CELLS THAT HAVE CELL PADDING
    # ON BOTH X AND Y AXES

    def _set_text_position(self, renderer):
        """Set text up so it is drawn in the right place."""
        bbox = self.get_window_extent(renderer)
        ha = self._text.get_horizontalalignment()
        va = self._text.get_verticalalignment()
        y_pad = self.PAD * bbox.width / bbox.height
        if ha == 'center': x = bbox.x0 + bbox.width / 2
        elif ha == 'left': x = bbox.x0 + bbox.width * self.PAD
        else: x = bbox.x0 + bbox.width * (1 - self.PAD)
        if va == 'center': y = bbox.y0 + bbox.height / 2
        elif va == 'top': y = (bbox.y0 + bbox.height * (1 - y_pad))
        else: y = (bbox.y0 + bbox.height * y_pad)
        self._text.set_position((x, y))


class Params(object):

    """Class for retrieving parameters from callable objects"""

    EXCLUDE = 'self', 'cls'

    @classmethod
    def get(cls, obj):
        """Returns the parameters for an object

        A mapping of seven parameter types is returned:
            args: a list of the parameter names
            varargs/varkw: the names of the * and ** parameters or None
            defaults: an n-tuple of the default values of the last n parameters
            kwonlyargs: a list of keyword-only parameter names
            kwonlydefaults: a mapping of names from kwonlyargs to defaults
            annotations: a mapping of parameter names to annotations
        """
        keys = ('args', 'varargs', 'varkw', 'defaults', 'kwonlyargs',
                'kwonlydefaults', 'annotations')
        params = dict(zip(keys, inspect.getfullargspec(obj)))
        return params

    @classmethod
    def getdefaults(cls, obj, ignore_none=False):
        """Returns the default parameters for an object, with the option
        to ignore defaults without a value
        """
        params = cls.get(obj)
        defaultparams = dict()
        args = params['args'] if params['args'] else list()
        defaults = list(params['defaults']) if params['defaults'] else list()
        kwonlyargs = params['kwonlyargs'] if params['kwonlyargs'] else list()
        kwonlydefaults = params['kwonlydefaults'] if params['kwonlydefaults']\
                         else dict()
        while len(args) > len(defaults): defaults.insert(0, None)
        defaultparams.update(dict(zip(args, defaults)))
        if kwonlyargs:
            defaultparams.update(dict((k, None) for k in kwonlyargs))
        if kwonlydefaults: defaultparams.update(kwonlydefaults)
        defaultparams = dict((k, v) for k, v in defaultparams.items()
                             if k not in cls.EXCLUDE)
        if ignore_none:
            defaultparams = dict((k, v) for k, v in defaultparams.items()
                                 if v is not None)
        return defaultparams

    @classmethod
    def getdefaultnames(cls, *objs):
        """Returns the names of default parameters for one or more objects
        """
        names = list()
        for obj in objs: names += list(cls.getdefaults(obj))
        return names


class Properties(object):

    """Class that contains necessary properties for objects in the module.

    --------------------------------------------------------------------------

    PROPERTIES REFERENCE:

        edgecolor: SEE COLOR REFERENCE
            The color of the lines for the box surrounding an object.
            
        facecolor: SEE COLOR REFERENCE
            The color of the fill for the box surrounding an object.
            
        fill: Boolean
            Specifies if the box surrounding an object should be filled.
            
        visible_edges: string
        
            Specifies the lines to be rendered of a box surrounding an
            object. Options consist of:
            
            - substring of 'BTRL' (Bottom, Top, Left, Right)
            - 'open' (no edges drawn)
            - 'closed' (all edges drawn)
            - 'horizontal' (top/bottom edges drawn)
            - 'vertical' (left/right edges drawn)
            
        linewidth: integer or float
            The width (in points) of the lines to be rendered of a box
            surrounding an object.
            
        linestyle: string or tuple of an integer or float and an even-length
                   tuple of integers or float values
        
            The style of the lines to be rendered of a box surrounding
            an object. Options consist of:
            - '-' or 'solid' (solid line)
            - '--' or 'dashed' (dashed line)
            - '-.' or 'dashdot' (dash-dot line)
            - ':' or 'dotted' (dotted line)
            - 'none', 'None', ' ', or '' (no line)
            - (offset, (line width, space width, ...)) (in points)
            
        hatch: string
        
            Specifies the type of hatching to use in the fill of a box
            surrounding an object, where the string is one or more specific
            characters, where longer combinations increase the density of
            the hatching pattern. Options consist of:
            
            - '/' (diagonal)
            - '\' (reversed diagonal)
            - '|' (vertical)
            - '-' (horizontal)
            - '+' (cross)
            - 'x' (diagonal cross)
            - 'o' (small circle)
            - 'O' (large circle)
            - '.' (dots)
            - '*' (stars)

            WARNING: Hatching patterns rely on the edge color of the
            containing box object and do not support alpha channel
            (transparency) use.
            
        capstyle: string
        
            Specifies the style of cap of the patterned lines to be rendered
            for a box containing an object. Options consist of:
            
            - 'butt' (flat termination at ends)
            - 'projecting' (squared termination at ends by line width)
            - 'round' (squared termination at ends by line width)
            
        joinstyle: string

            Specifies the style of line joining of the lines to be rendered
            for a box containing an object. Options consist of:
            
            - 'miter'
            - 'round'
            - 'bevel'

        padding: integer or float
            Specifies the minimum padding in points between the text of
            an object and the edges for the box containing the object.
            
        family: string
            The font family (system font name, 'serif', 'sans-serif',
            'cursive', 'fantasy', or 'monospace') for text.
        
        style: string
            The font style ('normal', 'italic', or 'oblique') for text.
            
        variant: string
            The font variant ('normal' or 'small-caps') for text.
            
        weight: integer, float, or string
        
            The font weight for text. Options consist of:
            
            - a numeric value from 0 to 1,000
            - 'ultralight'
            - 'light'
            - 'normal'
            - 'regular'
            - 'book'
            - 'medium'
            - 'roman'
            - 'semibold'
            - 'demibold'
            - 'demi'
            - 'bold'
            - 'heavy'
            - 'extra bold'
            - 'black'
            
        stretch: integer, float, or string
        
            The font stretch for text. Options consist of:
            
            - a numeric value from 0 to 1,000
            - 'ultra-condensed'
            - 'extra-condensed'
            - 'condensed'
            - 'semi-condensed'
            - 'normal'
            - 'semi-expanded'
            - 'expanded'
            - 'extra-expanded'
            - 'ultra-expanded'
            
        size: integer, float, or string
        
            The font size for text. Options consist of:
            
            - integer or float (in points)
            - 'xx-small'
            - 'x-small'
            - 'small'
            - 'medium'
            - 'large'
            - 'x-large'
            - 'xx-large'
            
        color: SEE COLOR REFERENCE
            The font color for text.
            
        verticalalignment: string
            The vertical alignment ('top', 'bottom', or 'center') of text.
            
        horizontalalignment: string
            The horizontal alignment ('left', 'right', or 'center') of text.
            
        multialignment: string
            The horizontal alignment ('left', 'right', or 'center') of
            multiline text.
            
        rotation: integer, float, or string
        
            The clockwise string-based rotation of text. Options consist of:
            
            - integer or float (in degrees)
            - 'vertical'
            - 'horizontal'
            
        linespacing: integer or float
            The spacing between lines of text in multiples of font size.
            
        wrap: Boolean
            Specifies if text wrapping should occur

        antialiased: Boolean
            Specifies whether to use antialiased rendering

    --------------------------------------------------------------------------

    COLOR REFERENCE:
    
        Supported colors include the following formats:
        
        - Case-insensitive X11 color name string (w/o spaces)
        - RGB or RGBA tuple of float values with an interval of [0, 1]
        - Case-insensitive hex RGB or RGBA string (e.g. '#0a0a0a' or
          '#0a0a0a0a')
        - String representation of a float value for grayscale (e.g. '0.4')
        - Single character shorthand notation for limited basic colors:
          - 'b' (blue)
          - 'g' (green)
          - 'r' (red)
          - 'c' (cyan)
          - 'm' (magenta)
          - 'y' (yellow)
          - 'k' (black)
          - 'w' (white)
          
    --------------------------------------------------------------------------
    """

    DIMENSIONS = ['x', 'y', 'xy', 'width', 'height']
    ALIGNMENTS = list(i for i in Params.getdefaultnames(BaseText)
                      if 'align' in i)
    FONT = Params.getdefaultnames(BaseFont)
    CELLSET = ['hatch']


Properties.EDGE = list(i for i in Params.getdefaultnames(BaseCell)
                       if 'edge' in i)
Properties.CELL = list(i for i in Params.getdefaultnames(BaseCell) if (i not
                       in Properties.EDGE and i != 'loc') or i in
                       Properties.DIMENSIONS)
Properties.EDGESET = list(
        i for i in Params.getdefaultnames(
            matplotlib.patches.Patch, matplotlib.patches.Rectangle)
        if i not in ['color'] + Properties.CELL + Properties.EDGE +
        Properties.CELLSET)
Properties.ALLCELL = Properties.CELL + Properties.CELLSET
Properties.ALLEDGE = Properties.EDGE + Properties.EDGESET
Properties.TEXT = list(i for i in Params.getdefaultnames(BaseText) if i not
                       in ['x', 'y'] + Properties.CELL + Properties.FONT)
Properties.ALL = list(sorted(set(Properties.ALLCELL + Properties.ALLEDGE +
                                 Properties.FONT + Properties.TEXT)))
Properties.FORMAT = list(i for i in Properties.ALL if i not in
                         Properties.DIMENSIONS + ['text'])
Properties.FORMATCELL = list(i for i in Properties.FORMAT if i in
                             Properties.CELL + Properties.CELLSET +
                             Properties.EDGE + Properties.EDGESET +
                             Properties.ALIGNMENTS)

Properties.NORESIZE = list(i for i in Properties.FORMAT if i not in
                           Properties.TEXT + Properties.FONT or i in
                           ['color', 'fontproperties'] + Properties.ALIGNMENTS)


class Page(object):

    """Class to create and manage a report page"""

    CONTAINER, LAYOUT, RENDERER = None, None, None
    SIZES = {'letter': (8.5, 11), 'legal': (8.5, 14), 'tabloid': (11, 17),
             'poster': (24, 30)}
    SIZES.update(dict(('{}-landscape'.format(k), tuple(reversed(v)))
                      for k, v in SIZES.items()
                      if tuple(v) != tuple(reversed(v))))
    WIDTH, HEIGHT = SIZES['letter']
    VERTICAL = ('top', 'bottom', 'center')
    HORIZONTAL = ('left', 'right', 'center')
    ALIGNMENTS = tuple('-'.join(i) for i in product(VERTICAL, HORIZONTAL)
                       if i[0] != i[1]) + ('center',)
    MARGINS = dict(zip(('top', 'bottom', 'left', 'right'), (1.0,) * 4))
    matplotlib.rcParams['figure.dpi'] = DPI = 300

    @classmethod
    def clear(cls):
        """Clears the current page"""
        if cls.CONTAINER is not None:
            cls.CONTAINER.clear()
            cls.LAYOUT = cls.CONTAINER.add_subplots()
            plot.subplots_adjust(0, 0, 1, 1, 0, 0)
        else: cls.create(width=cls.WIDTH, height=cls.HEIGHT)
        return

    @classmethod
    def create(cls, size=None, width=None, height=None):
        """Creates a blank page with optional size or width/height in inches;
        see Page.SIZES for size options
        """
        plot.close('all')
        cls.CONTAINER, cls.LAYOUT = plot.subplots()
        plot.subplots_adjust(0, 0, 1, 1, 0, 0)
        cls.CONTAINER.patch.set_visible(False)
        cls._setpagesize(size, width, height)
        cls.RENDERER = cls.CONTAINER.canvas.get_renderer()
        return

    @classmethod
    def getalignxy(cls, align=None, width=None, height=None):
        """Returns the Cartesiand x/y coordinates in inches for a specified
        alignment accounting for optional width/height in inches of an element
        inside the page margins
        """
        align = align if align in cls.ALIGNMENTS else cls.ALIGNMENTS[0]
        if any((width is None, height is None)): width, height = 0, 0
        if 'left' in align: x = Page.TypeArea.left()
        elif 'right' in align: x = Page.TypeArea.right() - width
        else: x = Page.TypeArea.x_center() - (width / 2)
        if 'top' in align: y = Page.TypeArea.top() - height
        elif 'bottom' in align: y = Page.TypeArea.bottom() 
        else: y = Page.TypeArea.y_center() - (height / 2)
        return x, y

    @classmethod
    def mergepdfs(cls, output_path, *pdf_paths):
        """Merges PDF files into a single file"""
        if pdf_paths:
            merger = PdfFileMerger()
            for pdf in pdf_paths: merger.append(pdf)
            merger.write(output_path)
            merger.close()
            for pdf in pdf_paths: os.remove(pdf)
        return

    @classmethod
    def refocus(cls):
        """Refocuses the page to the primary layout item"""
        try: plot.sca(Page.CONTAINER.axes[0])
        except: pass
        return

    @classmethod
    def savetopdf(cls, output_path):
        """Saves the page in a PDF format to a specified output path"""
        cls.saveto(output_path, 'pdf')
        return output_path

    @classmethod
    def saveto(cls, output_path, file_format=None):
        """Saves the page to the specified output path, with the option to
        specify png/pdf/svg/eps/ep format (or pdf otherwise)
        """
        if file_format not in ('png', 'pdf', 'svg', 'eps', 'ps'):
            file_format = 'pdf'
        bbox = matplotlib.transforms.Bbox.from_bounds(
            0.0, 0.0, cls.WIDTH, cls.HEIGHT)
        plot.savefig(output_path, bbox_inches=bbox, dpi=cls.DPI,
                     format=file_format)
        return output_path

    @classmethod
    def setdpi(cls, dpi=None):
        """Sets the dots per inch value for the page"""
        if not cls.LAYOUT:
            dpi = dpi if dpi else 300
            matplotlib.rcParams['figure.dpi'] = cls.DPI = abs(dpi)
            return
        error = 'DPI cannot be altered for a page that already exists. '\
                'Please set DPI prior to creating a new page.'
        raise Exception(error)
        return

    @classmethod
    def setmargins(cls, all_margins=None, top=None, bottom=None, left=None,
                   right=None):
        """Sets the page margins in inches"""
        if all_margins is not None:
            cls.MARGINS = dict(zip(cls.MARGINS.keys(), (all_margins,) * 4))
        if top: cls.MARGINS['top'] = abs(top)
        if bottom: cls.MARGINS['bottom'] = abs(bottom)
        if left: cls.MARGINS['left'] = abs(left)
        if right: cls.MARGINS['right'] = abs(right)
        return

    @classmethod
    def size(cls):
        """Returns the size (in inches) of the current page"""
        if Page.CONTAINER is not None:
            return '{} in. by {} in.'.format(
                Page.CONTAINER.bbox_inches.width,
                Page.CONTAINER.bbox_inches.height)
        return None
    
    @classmethod
    def _setpagesize(cls, size=None, width=None, height=None):
        """Sets the width/height in inches for page size"""
        if size is not None: cls.WIDTH, cls.HEIGHT = cls.SIZES[size]
        elif width and height: cls.WIDTH, cls.HEIGHT = width, height
        else: cls.WIDTH, cls.HEIGHT = cls.SIZES['letter']
        if cls.CONTAINER: cls.CONTAINER.set_size_inches(cls.WIDTH, cls.HEIGHT)
        return

    class TypeArea(object):

        """Class with methods that return information about the type area for
        the page
        """

        @classmethod
        def x_center(cls): return Page.WIDTH / 2
        @classmethod
        def left(cls): return Page.MARGINS['left']
        @classmethod
        def right(cls): return Page.WIDTH - Page.MARGINS['right']
        @classmethod
        def width(cls): return cls.right() - cls.left()

        @classmethod
        def y_center(cls): return Page.HEIGHT / 2
        @classmethod
        def top(cls): return Page.HEIGHT - Page.MARGINS['top']
        @classmethod
        def bottom(cls): return Page.MARGINS['bottom']
        @classmethod
        def height(cls): return cls.top() - cls.bottom()


class Image(object):
    
    """Class for adding images to a report page"""

    @classmethod
    def add(cls, image, width, height, align=None, rotation=None, alpha=None,
            grayscale=False, clip=False, x=None, y=None):
        """Adds an image to the page
        
        ----------------------------------------------------------------------

        PARAMETERS:
        
            image: string or PIL.Image object
                A path to an image or a PIL.Image object.
        
            width: integer or float
                The width of the image in inches.

            height: integer or float:
                The height of the image in inches.

            align (optional): string

                The alignment of the image relative to the type area for the
                page. If x and y are specified, align will be ignored. Options
                consist of:

                - 'top-left'
                - 'top-right'
                - 'top-center'
                - 'bottom-left'
                - 'bottom-right'
                - 'bottom-center'
                - 'center-left'
                - 'center-right'
                - 'center'

                DEFAULT: 'top-left'
                
            rotation (optional): integer or float
                The clockwise rotation of the image in degrees.

            alpha (optional): integer or flot within [0.0, 0.1]
                The alpha compositing (transparency) of the image, where 0.0
                is completely transparent and 1.0 is opaque. DEFAULT: 1.0

            grayscale (optional): Boolean
                Specifies if the image will be converted to grayscale.
                DEFAULT: False

            clip (optional): Boolean
                Specifies whether to clip the image to optimal dimensions to
                be within the bounds of the specified width/height.
                DEFAULT: False

            x (optional): integer or float
                The cartesian origin of the image on the x-axis of the page.
                If specified with y, will override align.
                
            y (optional): integer or float
                The cartesian origin of the image on the y-axis of the page.
                If specified with x, will override align.

        ----------------------------------------------------------------------
        """
        mode, blank = 'RGBA', 'white'
        alpha = alpha if alpha is not None else 1.0
        if isinstance(image, PIL.Image.Image): image = image.convert(mode)
        else: image = PIL.Image.open(image).convert(mode)
        if grayscale: image = image.convert('LA').convert(mode)
        imgwidth, imgheight = image.size
        if clip and imgwidth > imgheight: height *= imgheight / imgwidth
        elif clip and imgwidth < imgheight: width *= imgwidth / imgheight
        align = align if align in Page.ALIGNMENTS else Page.ALIGNMENTS[0]
        if rotation:
            image = image.rotate(360-rotation, expand=True)
            background = PIL.Image.new(mode, image.size, blank)
            image = PIL.Image.composite(image, background, image)
        if x is None or y is None: x, y = None, None
        if x is None and y is None:
            x, y = Page.getalignxy(align, width, height)
        bounds = [x/Page.WIDTH, y/Page.HEIGHT,
                  width/Page.WIDTH, height/Page.HEIGHT]
        imagebox = Page.CONTAINER.add_axes(bounds)
        imagebox.axis('off')
        imagebox.imshow(image, alpha=alpha)
        Page.refocus()
        return imagebox


class Text(object):

    """Class to add and evaluate text for a report page"""
    
    _DUMMYCELL = BaseCell((0.0, 0.0), 1, 1, text='', visible_edges='open')

    @classmethod
    def add(cls, value, line_number=None, **properties):
        """Adds text to the page

        ----------------------------------------------------------------------

        PARAMETERS:

            value: any value
                The value to add as text; multiline values are accepted.

            line_number (optional): integer
                Specifies the line number to add text, based on the font and
                font size. DEFAULT: 1

            properties (optional): keyword arguments
                Specifies the formatting properties for the text. See the help
                documentation for the 'Properties' class in this module.
        
        ----------------------------------------------------------------------
        """
        value = cls.format(value)
        if line_number is not None: value = '\n' * (line_number - 1) + value
        for k, v in zip(['verticalalignment', 'horizontalalignment'],
                        Page.ALIGNMENTS[0].split('-')):
            if k not in properties: properties[k] = v
        rendered = Table([[value]], breakrowvalues=True, padding=0,
                         visible_edges='open', wrap=True,
                         rowheight=Page.TypeArea.height(),
                         columnwidths=Page.TypeArea.width())
        return rendered

    @classmethod
    def addbox(cls, value, width, height, align=None, x=None, y=None,
               **properties):
        """Adds a text box to the page

        ----------------------------------------------------------------------

        PARAMETERS:

            value: any value
                The value to add as text; multiline values are accepted.

            width: integer or float
                The width of the text box in inches.

            height: integer or float:
                The height of the text box in inches.

            align (optional): string

                The alignment of the text box relative to the type area for
                the page. If x and y are specified, align will be ignored.
                Options consist of:

                - 'top-left'
                - 'top-right'
                - 'top-center'
                - 'bottom-left'
                - 'bottom-right'
                - 'bottom-center'
                - 'center-left'
                - 'center-right'
                - 'center'

                DEFAULT: 'top-left'

            x (optional): integer or float
                The cartesian origin of the text box on the x-axis of the page.
                If specified with y, will override align.
                
            y (optional): integer or float
                The cartesian origin of the text box on the y-axis of the page.
                If specified with x, will override align.

            properties (optional): keyword arguments
                Specifies the configuration and formatting properties for the
                text box. Accepted properties are those that affect both text
                and table cells. See the help documentation for the
                'Properties' class in this module.
        
        ----------------------------------------------------------------------
        """
        if align is None: align = 'center'
        for k, v in zip(['verticalalignment', 'horizontalalignment'],
                        ['center', 'center']):
            if k not in properties: properties[k] = v
        box = Table([[value]], width, height, align, x, y, height, width,
                    **properties)
        return box
    
    @classmethod
    def format(cls, value):
        """Formats a value for use as cell text"""
        if value:
            value = '{}'.format(value).expandtabs()
            value = value.replace('\r\n', '\n').replace('\n\r', '\n')
            return value
        return ''

    @classmethod
    def getblock(cls, cell):
        text_object = cls._gettextobj(cell)
        return

    @classmethod
    def getheightwidth(cls, cell):
        bbox = cls._gettextbbox(cell)
        return bbox.height / Page.DPI, bbox.width / Page.DPI

    @classmethod
    def getheight(cls, cell):
        """Returns the height of the text for a cell in inches"""
        return cls._gettextbbox(cell).height / Page.DPI

    @classmethod
    def getwidth(cls, cell):
        """Returns the width of the text for a cell in inches"""
        return cls._gettextbbox(cell).width / Page.DPI

    @classmethod
    def getwrapwidth(cls, width=None):
        """Calculates text wrapping length based on page or specified width
        in inches
        """
        units = Page.LAYOUT.bbox.width / Page.WIDTH
        if width is None: width = Page.TypeArea.width()
        return int(math.ceil((units * width)))

    @classmethod
    def _gettextbbox(cls, cell):
        """Internal class method to return the bounding box for the text
        in a cell
        """
        return cls._gettextobj(cell).get_window_extent(Page.RENDERER)

    @classmethod
    def _gettextobj(cls, cell):
        """Internal class method to return a text object representing the
        cell text
        """
        tester = copy.deepcopy(cls._DUMMYCELL)
        cellparams = dict((k, v) for k, v in cell._cellparams.items()
                          if k != 'text')
        textparams = cell._textparams
        textparams['text'] = cell._text
        tester.set(**cellparams)
        tester.set_text_props(**textparams)
        tester.PAD = cell._PAD
        tester.visible_edges = ''
        text_object = copy.copy(tester.get_text())
        text_object._renderer = Page.RENDERER
        text_object.figure = Page.CONTAINER
        if cell.get('wrap'):
            wrapwidthin = cell._width - (cell._padding * 2 / 72)
            wrapfunc = lambda: cls.getwrapwidth(wrapwidthin)
            text_object._get_wrap_line_width = wrapfunc
            text_object.set(text=text_object._get_wrapped_text())
        return text_object


class _BaseClass(object):

    """Base class for cells, rows, tables, and table-like objects

    This base class provides basic functionality for tables and contained
    objects that allow for use as a partial mapping object along with
    specific configuration for indexing, attribute storage, and use for
    both internal and external attribute operations.
    """

    def get(self, key):
        """Gets a property of the object"""
        return self.__getitem__(key)

    def set(self, **properties):
        """Sets one or more properties of the object"""
        self.update(properties)
        return
    
    def update(self, properties):
        """Updates the object with a mapping object of properties"""
        for k, v in properties.items(): self.__setitem__(k, v)
        return

    @property
    def allproperties(self):
        """Returns all of the properties for the object"""
        return dict((self._outkey(k), v) for k, v in self.__dict__.items())

    @property
    def properties(self):
        """Returns the properties for the object that are specified"""
        return dict((k, v) for k, v in self.allproperties.items()
                    if v is not None)

    def _build(self): return
    def _inkey(self, key): return '_{}'.format(key)
    def _outkey(self, key): return key[1:]

    def _getattrs(self, names):
        """Internal method to retrieve specific object attributes based on
        specified names
        """
        return dict((i, self.get(i)) for i in names if self.get(i) is not None)

    def _setattrs(self, properties):
        """Internal method to set internal object attributes based on
        acceptable properties
        """
        try:
            if self._null: return
        except: pass
        accept = set(Params.getdefaultnames(self.__class__) +
                     Properties.ALL + ['padding', 'null'])
        for k, v in properties.items():
            if k not in accept:
                error = 'Invalid parameter specified ({}={}) '\
                        'for {}.'.format(k, v, self.__class__)
                raise Exception(error)
        attrs = dict((self._inkey(i), properties.get(i)) for i
                      in accept)
        attrs = dict((k, v) for k, v in attrs.items() if v is not None)
        self.__dict__.update(attrs)
        return

    def _clearattrs(self, properties_or_names):
        """Internal method to clear attributes based on properties or names
        USE WITH EXTREME CAUTION
        """
        for i in properties_or_names: self.__dict__.pop(self._inkey(i), None)
        return

    @property
    def _cellparams(self): return self._getattrs(Properties.CELL)

    @property
    def _cellsetparams(self):
        params = self._getattrs(Properties.CELLSET)
        params['linewidth'] = 0
        return params

    @property
    def _edgeparams(self):
        params = self._getattrs(Properties.EDGE)
        params.update(dict(text='', fill=False))
        return params

    @property
    def _edgesetparams(self):
        params = self._getattrs(Properties.EDGESET)
        if 'capstyle' not in params: params['capstyle'] = 'projecting'
        return params

    @property
    def _textparams(self):
        textprops = self._getattrs(Properties.TEXT)
        fontprops = self._getattrs(Properties.FONT)
        if fontprops: textprops['fontproperties'] = BaseFont(**fontprops)
        return textprops

    @property
    def _text(self): return Text.format(self.get('value'))

    def __getitem__(self, key):
        if isinstance(key, int): return self._getindex(key)
        if key == 'text': return self._text
        inkey = self._inkey(key)
        if key == 'rotation' and type(self.__dict__.get(inkey))\
           in (int, float): return -self.__dict__.get(inkey)
        elif key == 'xy':
            return (self.__dict__.get(self._inkey('x')),
                    self.__dict__.get(self._inkey('y')))
        return self.__dict__.get(inkey)

    def __setitem__(self, key, value):
        try:
            if self._null: return
        except: pass
        accept = set(Params.getdefaultnames(self.__class__) +
                     Properties.ALL + ['padding', 'null'])
        if key not in accept:
            error = 'Invalid parameter specified ({}={}) '\
                    'for {}.'.format(key, value, self.__class__)
            raise Exception(error)
        value = copy.copy(value)
        if isinstance(key, int): return self._setindex(key, value)
        key = 'value' if key == 'text' else key
        inkey = self._inkey(key)
        if key == 'rotation' and type(value) in (int, float): value = -value
        elif key == 'xy':
            self._x, self._y = value
            return
        self.__dict__[inkey] = value
        self._build()
        return

    def _getindex(self, key): return
    def _setindex(self, key, value): return


class _Cell(_BaseClass):

    """Internal object for use as an individual cell in a table.

    --------------------------------------------------------------------------

    USAGE:

        cell = _Cell({value}, {x}, {y}, {width}, {height}, {columnspan},
                     {rowspan}, {**properties})

        NOTE: Not recommended for use outside of a Table object. If provided
        to a table object upon instantiation, an error will occur.
    
    PARAMETERS:

        value (optional): any value
            The value of the cell.

        x (optional): integer or float


        y (optional): integer or float


        width (optional): integer or float


        height (optional): integer or float
        
            
        columnspan (optional): integer
            The number of columns the cell will span in a row.
            
        rowspan (optional): integer
            The number of rows the cell will span in a column.
            
        properties (optional): keyword arguments
            Specifies the configuration and formatting properties for the
            cell. See the help documentation for the 'Properties' class in
            this module.

    --------------------------------------------------------------------------
    """

    def __init__(self, value=None, x=0.0, y=0.0, width=1.0, height=1.0,
                 rowspan=1, columnspan=1, **properties):
        self._value = value
        self._x, self._y = x, y
        self._width, self._height = width, height
        self._rowspan, self._columnspan = rowspan, columnspan
        self._row, self._column, self._edges = None, None, list()
        self._null = False
        self._setattrs(properties)
        if self.__dict__.get('_padding') is None: self._padding = 2.25
        self._rendered = None
        self._build()
        return

    def copy(self): return self.__copy__()

    def format(self, override_edges=False, **properties):
        """Formats the cell

        PARAMETERS:

            override_edges (optional): Boolean

                Specifies if edge formatting should override and replace
                preexisting edge formatting for table cells. DEFAULT: False

            properties (optional): keyword arguments
            
                Specifies the configuration and formatting properties for
                cells. See the help documentation for the 'Properties' class
                in this module.
        """
        if self._null: return
        if override_edges: self._edges = list()
        edgeparams = dict((k, v) for k, v in properties.items() if k in
                          Properties.ALLEDGE)
        if edgeparams and 'visible_edges' not in edgeparams:
            edgeparams['visible_edges'] = 'BTRL'
        cellparams = dict((k, v) for k, v in properties.items() if k not in
                          Properties.ALLEDGE)
        if cellparams: self.set(**cellparams)
        if edgeparams:
            edges = _Edges()
            edges.set(**edgeparams)
            self._edges.append(edges)
        return

    def _build(self):
        """Internal method to build the cell object"""
        if self._null: self._nullify()
        else: self._buildedges()
        return

    def _buildedges(self):
        """Internal method to build the initial edges object for the cell"""
        params = self._edgeparams
        params.update(self._edgesetparams)
        edges = _Edges()
        edges.set(**params)
        if not self._edges: self._edges.insert(0, edges)
        else: self._edges[0] = edges
        return

    def _mergevaluewith(self, cell):
        """Internal method to merge the value of another cell into this cell"""
        if self._valueisnumeric() and cell._valueisnumeric():
            merged = self._value + cell._value
        else:
            value_a = self._value if self._value is not None else ''
            value_b = cell._value if cell._value is not None else ''
            merged = '{}{}'.format(value_a, value_b)
        self.set(value=merged)
        cell.set(value=None)
        return

    def _nullify(self):
        """Internal method to set the cell object to a null cell"""
        self._value, self._fill, self._edges = None, None, list()
        for i in self.__dict__:
            if 'color' in i: self.__dict__[i] = '#FFFFFFFF'
        return

    def _render(self, table=None):
        """Internal method to create a renderable object to add to a table,
        with normalization to table proportions when a table object is
        specified
        """
        if self._null: return
        if table is not None:
            x, y = self._x / table._width, self._y / table._height
            w, h = self._width / table._width, self._height / table._height
        else:
            x, y, w, h = self._x, self._y, self._width, self._height
        params = self._cellparams
        params.update(dict(xy=(x, y), width=w, height=h))
        rendered = BaseCell(**params)
        rendered.set(**self._cellsetparams)
        rendered.set_text_props(**self._textparams)
        rendered.PAD = self._PAD
        if self.get('wrap'):
            wrapwidth = self._width - (self._padding * 2 / 72)
            renderedtext = rendered.get_text()
            wrapfunc = lambda: Text.getwrapwidth(wrapwidth)
            renderedtext._get_wrap_line_width = wrapfunc
        self._rendered = rendered
        if table:
            table._renderarea.add_patch(self._rendered)
            self._validate(self._rendered)
        return

    def _validate(self, rendered):
        """Internal method to validate cell dimensions and y-axis location"""
        txtobj = rendered.get_text()
        txtobj._renderer = Page.RENDERER
        if self.get('wrap'): txtobj.set(text=txtobj._get_wrapped_text())
        bbox = txtobj.get_window_extent(Page.RENDERER)
        pad = 2 * self._padding / 72
        txtheight, txtwidth = bbox.height / Page.DPI, bbox.width / Page.DPI
        padheight, padwidth = txtheight + pad, txtwidth + pad
        if padwidth > self._width:
            error = 'An error has occurred resulting from a cell {}. '\
                    'The cell value exceeds the cell width. Please '\
                    'revise settings in cells to prevent this error and '\
                    'proceed.'.format(self)
            raise Exception(error)
        if padheight > self._height:
            error = 'An error has occurred resulting from a cell {}. '\
                    'The cell value exceeds the cell height. Please '\
                    'revise settings in cells to prevent this error and '\
                    'proceed.'.format(self)
            raise Exception(error)
        if self._y < 0:
            error = 'An error has occurred resulting from a cell {}. '\
                    'Rendering of the cell is outside of the render area '\
                    'of the table. Please remove data from the table or '\
                    'set the table to break at rows or row values to '\
                    'prevent this error and proceed.'.format(cell)
            raise Exception(self)
        return
        
    def _valueisnumeric(self): return type(self._value) in (int, float)

    @property
    def _PAD(self): return self._padding / 72 / self._width

    def __copy__(self):
        cls = self.__class__
        new = cls.__new__(cls)
        new.__dict__.update(self.__dict__)
        new._edges = list(i.copy() for  i in self._edges)
        return new

    def __repr__(self):
        value = repr(self._value) if not self._null else 'NULL'
        return '<Cell({},{}): {}>'.format(
            self._row if self._row is not None else 'N/A',
            self._column if self._column is not None else 'N/A', value)


class _Edges(_BaseClass):

    """Internal object class to handle cell edges"""

    def __init__(self, **properties):
        properties = dict((k, v) for k, v in properties.items() if k in
                          Properties.ALLEDGE)
        self._setattrs(properties)
        if 'padding' not in properties: self.set(padding=2.25)
        self._row, self._column, self._rendered = None, None, None
        return

    def copy(self): return self.__copy__()

    def _render(self, cell, table=None):
        """Internal method to create a renderable object to add to a table,
        with normalization to table proportions when a table object is
        specified
        """
        if cell._null: return
        if table is not None:
            x, y = cell._x / table._width, cell._y / table._height
            w, h = cell._width / table._width, cell._height / table._height
        else:
            x, y, w, h = cell._x, cell._y, cell._width, cell._height
        params = self._edgeparams
        params.update(dict(xy=(x, y), width=w, height=h))
        rendered = BaseCell(**params)
        rendered.set(**self._edgesetparams)
        self._rendered = rendered
        if table: table._renderarea.add_patch(self._rendered)
        return

    def __copy__(self):
        cls = self.__class__
        new = cls.__new__(cls)
        new.__dict__.update(self.__dict__)
        return new


class _Row(_BaseClass):

    """Internal object for use as an individual row of cells in a table.

    --------------------------------------------------------------------------

    USAGE:

        row = _Row([cell_1, cell_2, ... cell_n], {**properties})

        NOTE: Not recommended for use outside of a Table object. If provided
        to a table object upon instantiation, an error will occur.

    PARAMETERS:

        cells: sequence of _Cell objects
            A sequence of _Cell objects that will constitute the cells of the
            row.
        
        properties: keyword arguments
            Specifies the configuration and formatting properties for the
            row. See the help documentation for the 'Properties' class in
            this module.

    --------------------------------------------------------------------------
    """

    def __init__(self, cells=None, **properties):
        self._cells = cells or list()
        self._setattrs(properties)
        self._index = None
        self._build()
        return

    def format(self, column_index_or_index_range=None, override_edges=False,
               **properties):
        """Formats cells within the row

        PARAMETERS:
        
            column_index_or_index_range (optional): integer, range object, or
                sequence of integers and/or range objects

                Specifies the indices of columns to be formatted. Integers
                represent individual columns and range objects represent groups
                of columns, where values represent indices. If not specified,
                all columns will be formatted.
                DEFAULT: None

            override_edges (optional): Boolean

                Specifies if edge formatting should override and replace
                preexisting edge formatting for table cells. DEFAULT: False

            properties (optional): keyword arguments
            
                Specifies the configuration and formatting properties for
                cells. See the help documentation for the 'Properties' class
                in this module.
        """
        columnindex = self._getformatindex(column_index_or_index_range)
        isedge = any(list(i for i in properties if i in Properties.ALLEDGE))
        if isedge and 'visible_edges' not in properties:
            properties['visible_edges'] = 'BTRL'
        for start, stop in columnindex:
            start = start if not callable(start) else start()
            stop = stop if not callable(stop) else stop()
            indexrange = range(start, stop)
            if isedge:
                edgeindex = self._getvisibleedgeindex(indexrange, **properties)
            for index in indexrange:
                if isedge:
                    visible = edgeindex.get(index)
                    if visible is not None:
                        properties['visible_edges'] = visible
                    else: properties.pop('visible_edges', None)
                try: trytest = self._cells[index]
                except: break
                self._cells[index].format(override_edges, **properties)
        properties = dict((k, v) for k, v in properties.items() if k not in
                          Properties.FORMATCELL)
        if properties: self._build()
        return

    def _getformatindex(self, column_index_or_index_range):
        """Returns callable start/stop values for use in index-based column
        formatting
        """
        base = column_index_or_index_range
        if base is not None and isinstance(base, range): base = [base,]
        elif base is not None and not isinstance(base, Sequence):
            base = [base,]
        elif base is None: base = [[0, -1],]
        index = list()
        for i in base:
            if isinstance(i, range):
                if i.step == 1: index.append([i.start, i.stop])
                else:
                    for ii in i: index.append([ii, ii + i.step])
            elif i == [0, -1]: index.append([0, lambda: len(self._cells)])
            elif isinstance(i, Sequence):
                start, stop = i
                index.append([start, stop])
            elif i == -1:
                index.append([lambda: len(self._cells) - 1,
                              lambda: len(self._cells)])
            elif isinstance(i, int) and i >= 0: index.append([i, i + 1])
        return index

    def copy(self): return self.__copy__()

    def _build(self):
        """Internal method to build the row object"""
        self._buildcells()
        self._setcolumnspans()
        return

    def _buildcells(self):
        """Internal method to build the cell objects within the row object"""
        for index, cell in enumerate(self._cells): cell._column = index
        return

    def _setcolumnspans(self):
        """Internal method to validate and build cell spanning across multiple
        columns in the row object
        """
        maincell, spanranges = None, self._getspanranges()
        for cell in self:
            if self._canspan(cell, spanranges): maincell = cell
            elif self._inspan(cell, spanranges):
                maincell._mergevaluewith(cell)
                cell._null = True
        return

    def _canspan(self, cell, spanranges):
        """Internal method to check if a cell can span multiple columns
        without overlapping another cell that spans multiple columns in the
        row object
        """
        column, span = cell._column, cell._columnspan
        if span > 1 and any((column == min(i) for i in spanranges)):
            return True
        elif span == 1: return False
        error = 'An error occurred in {}. Multiple cells spanning multiple '\
                'columns cannot overlap in a row. Please revise properties '\
                'in cells to prevent this error.'.format(repr(cell))
        raise Exception(error)
        return

    def _inspan(self, cell, spanranges):
        """Internal method to check if a cell is within a range of columns
        that will contain a spanned cell in the row object
        """
        return any(list(cell._column in i for i in spanranges))

    def _getspanranges(self):
        """Internal method to return the column index ranges in which cells
        will span within a row object
        """
        return list(range(i._column, i._column + i._columnspan) for i
                    in self._cells if i._columnspan > 1)

    def _getindex(self, index):
        """Internal method to return a copy of a cell at a specific column
        index within the row object
        """
        return self._cells[index].copy()

    def _getvisibleedgeindex(self, indexrange, **properties):
        """Returns the visible edges indexed for application to cells"""
        edgeindex = dict()
        visible = BaseCell._edge_aliases.get(properties.get(
            'visible_edges')) or properties.get('visible_edges')
        if visible is not None and len(indexrange) > 1:
            edgelist = list(visible.replace('L', '').replace('R', '')
                            for i in indexrange)
            edgelist[0] = visible.replace('R', '')
            edgelist[-1] = visible.replace('L', '')
            edgeindex = dict(zip(indexrange, edgelist))
        elif visible is not None and len(indexrange) == 1:
            edgeindex[indexrange[0]] = visible
        return edgeindex

    @property
    def _isnull(self):
        """Internal property that indicates if all cells within the row object
        are null
        """
        return all(tuple(i._null for i in self))
    
    def __copy__(self):
        cls = self.__class__
        new = cls.__new__(cls)
        new.__dict__.update(self.__dict__)
        new._cells = list(i.copy() for i in self._cells)
        return new

    def __iter__(self):
        for cell in self._cells: yield cell

    def __len__(self): return len(self._cells)

    def __repr__(self):
        return '<Row({}): {}>'.format(
            self._index if self._index is not None else 'N/A', self._cells)


class Table(_BaseClass):

    """Object for use in rendering tabular data.

    --------------------------------------------------------------------------

    USAGE:

        table = Table([row_1, row_2, ... row_n], {width}, {height}, {align},
                      {x}, {y}, {rowheight}, {columnwidths}, {scalecolumns},
                      {expandrows}, {breakrows}, {breakrowvalues},
                      {padrowstotableheight}, {delayrender}, {**properties})

        - or -

        table = Table([valuelist_1, valuelist_2, ... valuelist_n], {width},
                      {height}, {align}, {x}, {y}, {rowheight},
                      {columnwidths}, {scalecolumns}, {expandrows},
                      {breakrows}, {breakrowvalues}, {padrowstotableheight},
                      {delayrender}, {**properties})
    
    PARAMETERS:

        rows_or_row_values: sequence of value sequences
            A sequence of value sequences that will constitute the table.
        
        x (optional): integer or float
            The cartesian origin of the table on the x-axis of the page. If
            specified with y, will override align.
            DEFAULT: x-axis origin of type area for page
            
        y (optional): integer or float
            The cartesian origin of the table on the y-axis of the page. If
            specified with x, will override align.
            DEFAULT: y-axis origin of type area for page

        width (optional): integer or float
            The width of the table in inches.
            DEFAULT: width of type area for page
            
        height (optional): integer or float
            The height of the table in inches.
            DEFAULT: height of type area for page

        align (optional): string

            The alignment of the table relative to the type area for the page.
            Alignment considers the entire table height and width, not just
            where table cells are rendered. If x and y are specified, align
            will be ignored. Options consist of:

            - 'top-left'
            - 'top-right'
            - 'top-center'
            - 'bottom-left'
            - 'bottom-right'
            - 'bottom-center'
            - 'center-left'
            - 'center-right'
            - 'center'

            DEFAULT: 'top-left'

        rowheight (optional): integer or float
            The height in inches of all rows in the table. DEFAULT: 0.25

        columnwidths (optional): integer, float, or sequence of integer/float
                                 values
            The width or sequence of widths in inches for all columns in the
            table. If specifying a sequence of widths, the number of widths
            must at least correspond to the number of columns in all rows or
            an error will occur. If specifying a single width, all columns
            will be assigned that width. DEFAULT: 0.75

        scalecolumns (optional): Boolean
            Specifies if columns should Ebe automatically scaled to fill the
            width of the table. DEFAULT: False

        expandrows (optional): Boolean
            Specifies if rows should be automatically expanded (in increments
            of row height) to contain wrapped text. NOTE: Cells that contain
            wrapped text must have wrap settings set to True or an error will
            occur. Cells containing appropriate line breaks do not require the
            wrap setting if the column widths accommodate text lines.
            DEFAULT: False

        breakrows (optional): Boolean
            Specifies if the table should break when the table rows reach
            the maximum height. If specified, remaining rows will be stored
            until the table is advanced to the next page. DEFAULT: False

        breakrowvalues (optional): Boolean
            Specifies if the table should break a row and its cell values
            when the table rows reach the maximum height. If specified, the
            top portion of the row will be rendered while the bottom portion
            and the remaining rows will be stored until the table is advanced
            to the next page. DEFAULT: False

        padrowstotableheight (optional): Boolean
            Specifies if the table should pad rows with additional empty rows
            to fill the table to the table height. The empty rows will be
            copies of the initial last row of the table. DEFAULT: False

        delayrender (optional): Boolean
            Specifies if the table should delay rendering after creation.
            This enables higher efficiency in changing the table after it
            has been created, especially where performing multiple changes.
            If this parameter is specified, the render method must be used
            to render the table.
            DEFAULT: False
        
        properties: keyword arguments
            Specifies the configuration and formatting properties for the
            row. See the help documentation for the 'Properties' class in
            this module.

    --------------------------------------------------------------------------
    """

    def __init__(self, rows_or_row_values, width=None, height=None,
                 align=None, x=None, y=None, rowheight=None,
                 columnwidths=None, scalecolumns=None, expandrows=None,
                 breakrows=None, breakrowvalues=None,
                 padrowstotableheight=None, delayrender=None, **properties):
        self._rows = rows_or_row_values or list()
        self._width = width
        self._height = height
        self._align = align
        self._x = x
        self._y = y
        self._rowheight = rowheight or 0.25
        self._columnwidths = columnwidths or 0.75
        self._scalecolumns = True if scalecolumns else False
        self._expandrows = True if expandrows else False
        self._breakrows = True if breakrows else False
        self._breakrowvalues = True if breakrowvalues else False
        self._padrowstotableheight = True if padrowstotableheight else False
        self._delayrender = True if delayrender else False
        self._setattrs(properties)
        self._rowsbuilt = False
        self._overflow = list()
        self._multipageformats = list()        
        self._tablepages = None
        self._renderarea = None
        self._build()
        return

    def format(self, row_index_or_index_range=None,
               column_index_or_index_range=None, multipage=False,
               override_edges=False, **properties):
        """Formats cells within the table

        PARAMETERS:

            row_index_or_index_range (optional): integer, range object, or
                sequence of integers and/or range objects

                Specifies the indices of rows to be formatted. Integers
                represent individual rows and range objects represent groups
                of rows, where values represent indices. If not specified,
                all rows will be formatted.
                DEFAULT: None

            column_index_or_index_range (optional): integer, range object, or
                sequence of integers and/or range objects

                Specifies the indices of columns to be formatted. Integers
                represent individual columns and range objects represent groups
                of columns, where values represent indices. If not specified,
                all columns will be formatted.
                DEFAULT: None

            multipage (optional): Boolean

                Specifies if the formatting should be applied on a page-by-
                page basis (True) or considering all rows. DEFAULT: False

            override_edges (optional): Boolean

                Specifies if edge formatting should override and replace
                preexisting edge formatting for table cells. DEFAULT: False

            properties (optional): keyword arguments
            
                Specifies the configuration and formatting properties for
                cells. See the help documentation for the 'Properties' class
                in this module.
        """
        self._format(
            row_index_or_index_range, column_index_or_index_range, multipage,
            override_edges, render=True, **properties)
        return    

    def _format(self, row_index_or_index_range=None,
               column_index_or_index_range=None, multipage=False,
               override_edges=False, current_page_only=False, render=False,
                **properties):
        """Internal method for formatting table cells"""
        rows = self._rows if current_page_only else self._allrows
        rowindex = self._getformatindex(row_index_or_index_range, multipage)
        columnindex = column_index_or_index_range
        isedge = any(list(i for i in properties if i in Properties.ALLEDGE))
        if isedge and 'visible_edges' not in properties:
            properties['visible_edges'] = 'BTRL'
        if multipage:
            params = dict(row_index_or_index_range=rowindex,
                          column_index_or_index_range=columnindex,
                          override_edges=override_edges,
                          current_page_only=True)
            params.update(properties)
            self._multipageformats.append(params)
            if not self._delayrender: self._render()
            return
        for start, stop in rowindex:
            start = start if not callable(start) else start()
            stop = stop if not callable(stop) else stop()
            indexrange = range(start, stop)
            if isedge:
                edgeindex = self._getvisibleedgeindex(indexrange, **properties)
            for index in indexrange:
                if isedge:
                    visible = edgeindex.get(index)
                    if visible is not None:
                        properties['visible_edges'] = visible
                    else: properties.pop('visible_edges', None)
                try: trytest = rows[index]
                except: break
                rows[index].format(
                    columnindex, override_edges, **properties)
        properties = dict((k, v) for k, v in properties.items() if k not in
                          Properties.NORESIZE)
        if properties: self._build()
        return

    def ismultipage(self):
        """Returns true if the table spans multiple pages"""
        return True if self._overflow else False

    def nextpage(self):
        """Advances the table to the next page"""
        self._rows, self._overflow = self._overflow, list()
        if self._rows: self._build()
        return

    def remove(self):
        """Removes the table from the page"""
        if self._renderarea is not None: self._renderarea.remove()
        return

    def render(self, removedelay=False):
        """Renders the table, with the option to remove the render delay
        and resume automatic rendering for the table
        """
        if removedelay: self._delayrender = False
        self._render()
        return

    @property
    def pages(self):
        """Returns the number of pages for the table"""
        return self._tablepages

    def _getformatindex(self, row_index_or_index_range, multipage=False):
        """Returns callable start/stop values for use in index-based row
        alteration
        """
        base = row_index_or_index_range
        if base is not None and isinstance(base, range): base = [base,]
        elif base is not None and not isinstance(base, Sequence):
            base = [base,]
        elif base is None: base = [[0, -1],]
        index = list()
        for i in base:
            if isinstance(i, range):
                if i.step == 1: index.append([i.start, i.stop])
                else:
                    for ii in i: index.append([ii, ii + i.step])
            elif i == [0, -1] and multipage:
                index.append([0, lambda: len(self._rows)])
            elif i == [0, -1]: index.append([0, lambda: len(self._allrows)])
            elif isinstance(i, Sequence):
                start, stop = i
                index.append([start, stop])
            elif i == -1 and multipage:
                index.append([lambda: len(self._rows) - 1,
                              lambda: len(self._rows)])
            elif i == -1:
                index.append([lambda: len(self._allrows) - 1,
                              lambda: len(self._allrows)])
            elif isinstance(i, int) and i >= 0: index.append([i, i + 1])
        return index

    def _getvisibleedgeindex(self, indexrange, **properties):
        """Returns the visible edges indexed for application to cells"""
        edgeindex = dict()
        visible = BaseCell._edge_aliases.get(properties.get(
            'visible_edges')) or properties.get('visible_edges')
        if visible is not None and len(indexrange) > 1:
            edgelist = list(visible.replace('B', '').replace('T', '')
                            for i in indexrange)
            edgelist[0] = visible.replace('B', '')
            edgelist[-1] = visible.replace('T', '')
            edgeindex = dict(zip(indexrange, edgelist))
        elif visible is not None and len(indexrange) == 1:
            edgeindex[indexrange[0]] = visible
        return edgeindex

    def _build(self):
        """Internal method to build the table object"""
        self._buildforstructure()
        self._buildforsize()
        self._buildforpage()
        return

    def _buildforstructure(self):
        """Internal method that builds the table structure"""
        self._remergedata()
        self._buildrows()
        self._indexrows()
        return

    def _buildforsize(self):
        """Internal method that builds size aspects of the table"""
        self._validatetablesize()
        self._validatecolumnsizes()
        self._setrowspans()
        self._setrowsizes()
        self._setcolumnsizes()
        self._setrowexpansion()
        return

    def _buildforpage(self):
        """Internal method that builds page-related aspects of the table"""
        self._setpositions()
        self._padrows()
        self._setbreak()
        self._settablepages()
        if not self._delayrender: self._render()
        return

    def _remergedata(self):
        """Internal method to remerge table row and overflow rows"""
        if self._rows and self._overflow:
            row_a = self._rows[-1]
            row_b = self._overflow[0]
            if row_a._index == row_b._index:
                self._mergebrokenrows(row_a, row_b)
                self._rows += self._overflow[1:]
            else: self._rows += self._overflow
            self._overflow.clear()
        return

    def _mergebrokenrows(self, row_a, row_b):
        """Internal method to merge two rows that were previously one row
        prior to row value breaking
        """
        for a, b, in zip(row_a, row_b):
            a._height = self._rowheight
            if a._value and b._value: a._value = '\n'.join((a._text, b._text))
            a._y -= b._y
            a.set(height=a._height, y=a._y)
        return

    def _buildrows(self):
        """Internal method to build row objects within the table object"""
        if self._rowsbuilt: return
        built = list()
        formatparams = self._getattrs(Properties.FORMAT)
        self._rows = list(_Row(list(_Cell(value, **formatparams) for value in
                                    row)) for row in self._rows)
        self._clearattrs(formatparams)
        self._rowsbuilt = True
        return

    def _indexrows(self):
        """Internal method to index rows and cells within the table object"""
        for index, row in enumerate(self):
            row._index = index
            for cell in row: cell._row = index
        return  

    def _getindex(self, index):
        """Internal method to return a copy of a row at a specific row
        index within the table object
        """
        return self._rows[index].copy()

    def _validatetablesize(self):
        """Internal method to validate and set the table height/width to
        the page type area if not otherwise specified
        """
        if Page.CONTAINER is None: Page.create()
        if self._width is None: self._width = Page.TypeArea.width()
        if self._height is None: self._height = Page.TypeArea.height()
        return

    def _validatecolumnsizes(self):
        """Internal method to validate and set column widths where scaling
        is specified
        """
        if isinstance(self._columnwidths, Sequence) and self._scalecolumns:
            basewidth = sum(self._columnwidths)
            if basewidth != self._width:
                self._columnwidths = list(self._width * i / basewidth for i
                                          in self._columnwidths)
                return
        elif self._scalecolumns:
            error = 'Columns cannot be scaled to the table width unless '\
                    'individual column widths are specified.'
            raise Exception(error)
        return

    def _setcolumnsizes(self):
        """Internal method to set the column sizes for cells in the table
        object
        """
        widths = self._columnwidths
        multisize = isinstance(widths, Sequence)
        for row in self:
            for cell in row:
                if multisize and len(row) <= len(widths):
                    start = cell._column
                    stop = cell._column + cell._columnspan
                    width = sum(widths[start:stop])
                    cell._width = width
                elif not multisize:
                    cell._width = widths * cell._columnspan
                else:
                    error = 'A mismatch between specified column widths and '\
                            'row columns has occurred ({} column widths for '\
                            '{} columns. Please specify column widths to '\
                            'account for all columns.'\
                            ''.format(len(widths), len(row))
                    raise Exception(error)
        return

    def _setrowsizes(self):
        """Internal method to set the row sizes for cells in the table
        object
        """
        for row in self:
            for cell in row: cell._height = self._rowheight * cell._rowspan
        return

    def _setrowspans(self):
        """Internal method to validate and build cell spanning across multiple
        rows in the table object
        """
        for columnindex, rowranges in self._getspanranges().items():
            for rowrange in rowranges:
                maincell = None
                for rowindex in rowrange:
                    cell = self._getcell(rowindex, columnindex)
                    if cell._rowspan > 1: maincell = cell
                    else:
                        if maincell._columnspan > 1:
                            cell._columnspan = maincell._columnspan
                            self._rows[rowindex]._build()
                            cell = self._getcell(rowindex, columnindex)
                        maincell._mergevaluewith(cell)
                        cell._null = True
        return

    def _getspanranges(self):
        """Internal method to return the row index ranges in which cells
        will span within a row object
        """
        spanranges = dict()
        for row in self:
            for cell in row:
                column = cell._column
                if all((self._canspan(cell, spanranges),
                        not self._inspan(cell, spanranges))):
                    if not spanranges.get(column): spanranges[column] = list()
                    stop = cell._row + cell._rowspan
                    if stop > len(self._rows): stop = len(self._rows)
                    spanranges[column].append(range(cell._row, stop))
        return spanranges

    def _canspan(self, cell, spanranges):
        """Internal method to check if a cell can span multiple rows
        without overlapping another cell that spans multiple rows in the
        table object
        """
        column, row, rowspan = cell._column, cell._row, cell._rowspan
        if column not in spanranges and rowspan > 1: return True
        elif column in spanranges and rowspan > 1:
            if not any(list(row in i for i in spanranges[column])):
                return True
        elif rowspan == 1: return False
        error = 'An error has occurred resulting from a cell {}. '\
                'Multiple cells spanning multiple rows cannot overlap in '\
                'a column. Please revise settings in cells to prevent this '\
                'overlap and proceed.'.format(cell)
        raise Exception(error)
        return False

    def _inspan(self, cell, spanranges):
        """Internal method to check if a cell is within a range of rows
        that will contain a spanned cell in the table object
        """
        if cell._column in spanranges:
            return any(list(cell._row in i for i in spanranges[cell._column]))
        return False
        
    def _getcell(self, rowindex, columnindex):
        """Internal method to retrieve a table cell at the row and column
        indices; DOES NOT PROVIDE CELL/ROW/TABLE BUILD SUPPORT WITH RETRIEVAL
        """
        return self._rows[rowindex]._cells[columnindex]

    def _setpositions(self):
        """Internal method to set the positions of cells in the table based
        on cell attributes
        """
        xpos, ypos = 0, list(self._height for i in range(len(self[0])))
        for row in self:
            for index, cell in enumerate(row):
                if not cell._null:
                    cell._x = xpos
                    cell._y = ypos[index] - cell._height
                    for i in range(index, index + cell._columnspan):
                        ypos[i] -= cell._height
                xpos += self._getcolumnwidth(cell)
            xpos = 0
        return

    def _getcolumnwidth(self, cell):
        """Internal method to return the column width for a cell"""
        if isinstance(self._columnwidths, Sequence):
            return self._columnwidths[cell._column]
        return self._columnwidths

    def _setrowexpansion(self):
        """Internal method to expand row sizes to accommodate contained text
        in each cell, including with accommodation of cells spanning multiple
        rows
        """
        if not self._expandrows: return
        expandindex = dict()
        for row in self:
            index = row._index
            ratio = max(self._getexpansionratio(i) for i in row)
            if ratio: expandindex[index] = ratio
        self._expandunspannedcells(expandindex)
        self._expandspannedcells(expandindex) 
        return

    def _expandunspannedcells(self, expandindex):
        """Internal method to expand row sizes for cells that do not span
        multiple rows
        """
        for index, ratio in expandindex.items():
            expansion = ratio * self._rowheight
            for row in self:
                for cell in row:
                    if not cell._rowspan > 1:
                        if row._index == index:
                            cell._height += expansion
                            self._checkexpansion(cell)
        return

    def _expandspannedcells(self, expandindex):
        """Internal method to expand row sizes for cells that only span
        multiple rows
        """
        for index, ratio in expandindex.items():
            expansion = ratio * self._rowheight
            for row in self:
                for cell in row:
                    if cell._rowspan > 1 and row._index == index:
                        added = sum(self[i][cell._column]._height for i in
                                    range(index + 1, index + cell._rowspan))
                        newheight = expansion + self._rowheight + added
                        cell._height = newheight
                        self._checkexpansion(cell)
        return

    def _checkexpansion(self, cell):
        """Internal method to validate that a cell height does not exceed
        the total height of the table
        """
        if cell._height <= self._height: return
        error = 'An error has occurred resulting from a cell {}. '\
                'The cell height exceeds the table height for the page. '\
                'Please revise settings in cells to prevent this '\
                'error and proceed.'.format(cell)
        raise Exception(error)
        return

    def _getexpansionratio(self, cell):
        """Internal method to return the expansion ratio of row heights to
        determine the appropriate height for the text in a cell
        """
        base = math.ceil(cell._height / self._rowheight)
        ratio = math.ceil(self._getpaddedtextheight(cell) / self._rowheight)
        return ratio - base if ratio > base else 0

    def _getpaddedtextheight(self, cell):
        """Internal method to returns the height of the text plus padding
        for a cell in inches
        """
        textheight = Text.getheight(cell)
        padding = 2 * cell._padding / 72
        return padding + textheight

    def _getpaddedtextheightwidth(self, cell):
        """Internal method to returns the height and width of the text plus
        padding for a cell in inches
        """
        textheight, textwidth = Text.getheightwidth(cell)
        padding = 2 * cell._padding / 72
        return textheight + padding, textwidth + padding

    def _padrows(self):
        """Internal method to pad the table with blank rows to fill the
        render area height of the table
        """
        if not self._padrowstotableheight: return
        totalheight = self._gettableheight()
        blank = self[-1].copy()
        for cell in blank: cell.set(value='')
        while totalheight / self._height != math.ceil(
            totalheight / self._height):
            newheight = totalheight + self._rowheight
            if totalheight + self._rowheight / self._height > math.ceil(
                totalheight + self._rowheight / self._height):
                break
            else:
                totalheight += self._rowheight
                self._rows.append(blank.copy())
        self._indexrows()
        return

    def _settablepages(self):
        """Internal method to determine total number of pages for the table
        """
        height = self._gettableheight()
        self._tablepages = int(math.ceil(height / self._height))
        return

    def _gettableheight(self):
        """Internal method to get the height of the entire table in inches
        regardless of pages
        """
        height = 0
        for row in self:
            rowheights = list(i._height for i in row if i._rowspan == 1)
            if row._isnull or not rowheights: height += self._rowheight
            else: height += max(rowheights)
        return height

    def _setbreak(self):
        """Internal method to set the break point of the table rows to fit
        within the table height, breaking either at a full row or cell
        values
        """
        if not self._breakrowvalues and not self._breakrows: return
        index = None
        for row in self:
            baseline = min(i._y for i in row)
            if baseline < 0:
                index = row._index
                break
        if index is not None:
            self._overflow = list(i for i in self._overflow if i
                                  not in self._rows)
        if index is not None and self._breakrowvalues:
            lastrow, breakrow = self._getrowvaluebreak(self._rows[index])
            if lastrow is not None:
                self._overflow = [breakrow] + self._rows[index + 1:] + \
                                 self._overflow
                if all((i._height for i in lastrow)):
                    self._rows = self._rows[:index] + [lastrow]
                else: self._rows = self._rows[:index]
            else:
                self._rows, self._overflow = (
                    self._rows[:index], self._rows[index:] + self._overflow)
        elif index is not None:
            self._rows, self._overflow = (self._rows[:index],
                                          self._rows[index:] + self._overflow)
        return

    def _getrowvaluebreak(self, row):
        """Internal method to break row values into two rows, the first of
        which will be the last row of the table on the current page and the
        second will be the first row of the table on the next page
        """
        adjrow, newrow = row.copy(), row.copy()
        for index in range(len(row)):
            cell = row._cells[index]
            if not cell._null:
                adjcell, newcell = self._getcellvaluebreak(cell)
                adjrow._cells[index], newrow._cells[index] = adjcell, newcell
        if all((i is None or i._value is None for i in adjrow)):
            return None, newrow
        return adjrow, newrow

    def _getcellvaluebreak(self, cell):
        """Internal method to break a cell value into two values, the first of
        which will be in the last row of the table on the current page and the
        second will be in the first row of the table on the next page
        """
        if cell._height == self._rowheight: return None, cell
        if cell._rowspan > 1:
            error = 'An error has occurred resulting from a cell {}. '\
                    'Table rendering does not support splitting a cell '\
                    'across pages where the cell spans multiple rows. '\
                    'Please revise settings in cells to prevent this '\
                    'error and proceed.'.format(cell)
            raise Exception(error)
        adjcell, newcell = cell.copy(), cell.copy()
        while adjcell._y < 0:
            adjcell._y += self._rowheight
            adjcell._height -= self._rowheight
        newcell._height = newcell._height - adjcell._height
        newcell._y = newcell._y + adjcell._height
        if adjcell._rowspan > 1:
            adjcell._rowspan = int(math.ceil(adjcell._height/self._rowheight))
            newcell._rowspan = newcell._rowspan - adjcell._rowspan or 1
        if cell._value:
            celltext = Text._gettextobj(cell)
            if cell.get('wrap'): textblock = celltext._get_wrapped_text()
            else: textblock = celltext.get_text()
            if adjcell._height == 0:
                newcell._value, adjcell._value = adjcell._value, ''
            else:
                adjlines = textblock.split('\n')
                newcell._value = None
                newlines = list()
                while self._getpaddedtextheight(adjcell) > adjcell._height:
                    newlines.insert(0, adjlines[-1])
                    adjlines = adjlines[:-1]
                    adjcell.set(value='\n'.join(adjlines))
                newcell.set(value='\n'.join(newlines))
        return adjcell, newcell

    def _createrenderarea(self):
        """Internal method to create the render area for the table"""
        self.remove()
        if None in (self._x, self._y): x, y = None, None
        else: x, y = self._x, self._y
        if x is None and y is None:
            x, y = Page.getalignxy(self._align, self._width, self._height)
        bounds = [x / Page.WIDTH, y / Page.HEIGHT, self._width / Page.WIDTH,
                  self._height / Page.HEIGHT]
        self._renderarea = Page.CONTAINER.add_axes(bounds)
        self._renderarea.axis('off')
        self._x, self._y = x, y
        return

    def _render(self):
        """Internal method to render table cells in the render area"""
        self._createrenderarea()
        self._renderarea.clear()
        self._renderarea.axis('off')
        for i in self._multipageformats: self._format(**i)
        for row in self._rows:
            for cell in row: self._rendercell(cell)
        for row in self._rows:
            for cell in row: self._renderedges(cell)
        Page.refocus()
        return

    def _rendercell(self, cell):
        """Internal method to validate and render a table cell in the
        render area
        """
        if not cell._null: cell._render(table=self)
        return

    def _renderedges(self, cell):
        """Internal method to render the edges for a table cell in the
        render area
        """
        if cell._null: return
        for edges in cell._edges: edges._render(cell=cell, table=self)
        return

    def _getpaddedtextwidth(self, cell):
        """Internal method to return the width of the text plus cell padding
        for a cell in inches
        """
        padding = 2 * cell._padding / 72
        textwidth = Text.getwidth(cell)
        return padding + textwidth

    @property
    def _allrows(self):
        """Internal method to return all table rows, including overflow rows
        """
        return self._rows + self._overflow

    def __iter__(self):
        for row in self._rows + self._overflow: yield row

    def __len__(self): return len(self._rows) + len(self._overflow)
    
