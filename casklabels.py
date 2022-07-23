"""Tool for creating cask labels

The tool is driven from three CSV files.
- The label file defines the shape, size and pitch of the labels on the sheet
- The layout file defines what gets printed and where
- The data file contains the data that will be printed, one row per label

"""
import csv
import optparse
import logging
import sys
import os

from reportlab.pdfgen.canvas import Canvas
from reportlab.lib import pagesizes, units, styles, colors
from reportlab.platypus.frames import Frame
from reportlab.platypus.paragraph import Paragraph

# Do not use before logging config is called
logger = logging.getLogger("casklabels")

LOG_FORMAT = "%(asctime)s %(levelname)s %(message)s"

def find_page_size(name):
    """Get a page size object from its name

    """
    if not hasattr(pagesizes, name) or len(getattr(pagesizes, name)) != 2:
        logger.error("Page size '%s' not known", name)
        return None

    return getattr(pagesizes, name)


class LabelType(object):
    """Represents the type of label being used, including all dimensions

    """

    NAME_COLUMN = "Name"
    PAGE_SIZE_COLUMMN = "Page size"
    LEFT_COLUMN = "Left"
    HOR_PITCH_COLUMN = "Horizontal Pitch"
    WIDTH_COLUMN = "Width"
    COLUMNS_COLUMN = "Columns"
    TOP_COLUMN = "Top"
    VER_PITCH_COLUMN = "Vertial Pitch"
    HEIGHT_COLUMN = "Height"
    ROWS_COLUMN = "Rows"

    COLUMNS = [NAME_COLUMN, PAGE_SIZE_COLUMMN,
               LEFT_COLUMN, HOR_PITCH_COLUMN, WIDTH_COLUMN, COLUMNS_COLUMN,
               TOP_COLUMN, VER_PITCH_COLUMN, HEIGHT_COLUMN, ROWS_COLUMN,]


    @classmethod
    def read_file(cls, csvfile):
        """Reads the label definition file and creates one instance per row

        csvfile - name of file to be read

        returns dictionary from label type names to LabelType instances

        The first label type in the file is also keyed on None so that it can
        be used as a default type

        """
        with open(csvfile, "r") as stream:
            reader = csv.reader(stream)
            header = next(reader)
            if cls.check_header(header) > 0:
                return {}

            label_list = [cls(dict(zip(header, row)))
                          for row in reader]

            label_dict = dict([(lt.name, lt)
                               for lt in label_list])
            if len(label_list) > 0:
                label_dict[None] = label_list[0]

        return label_dict

    @classmethod
    def check_header(cls, header):
        """Check tha the label type file's header is valid

        """
        errors = 0
        for column_name in cls.COLUMNS:
            if column_name not in header:
                logger.error("Label file must have a '%s' column",
                             column_name)
                errors += 1
        return errors

    @classmethod
    def create_file(cls, csvfile):
        """Create an empty label definition file, just the headers are included

        """
        if os.path.exists(csvfile):
            logger.error("Label file '%s' exists, not overwriting",
                         csvfile)
            return

        with open(csvfile, 'w') as stream:
            writer = csv.writer(stream)
            writer.writerow(cls.COLUMNS)

    def __init__(self, parameters):
        """Create a LabelType instance for the given parameters

        """
        self.parameters = parameters
        self.name = parameters[self.NAME_COLUMN]
        self.page_size = find_page_size(parameters[self.PAGE_SIZE_COLUMMN])

        self.left = float(parameters[self.LEFT_COLUMN]) * units.mm
        self.hor_pitch = float(parameters[self.HOR_PITCH_COLUMN]) * units.mm
        self.width = float(parameters[self.WIDTH_COLUMN]) * units.mm
        self.columns = int(parameters[self.COLUMNS_COLUMN])

        self.top = float(parameters[self.TOP_COLUMN]) * units.mm
        self.ver_pitch = float(parameters[self.VER_PITCH_COLUMN]) * units.mm
        self.height = float(parameters[self.HEIGHT_COLUMN]) * units.mm
        self.rows = int(parameters[self.ROWS_COLUMN])

        self.bottom = (self.page_size[1] - self.top -
                       self.rows * self.ver_pitch)
        self.labels_per_page = self.rows * self.columns



    def start_label_run(self, canvas):
        """Prepare the canvas for a run of labels

        """
        if self.page_size is not None:
            canvas.setPageSize(self.page_size)

    def start_label(self, canvas, label_number):
        """Prepare the canvas for the next label

        label_number - zero based number of the label to be started

        """
        # Rows need to be counted from the bottom
        row = self.rows - 1 - (label_number // self.columns) % self.rows
        column = label_number % self.columns

        if label_number > 0 and (label_number % self.labels_per_page) == 0:
            canvas.showPage()

        canvas.saveState()

        canvas.translate(self.left + column * self.hor_pitch,
                         self.bottom + row * self.ver_pitch)
        #canvas.rect(0, 0, self.width, self.height)



    def end_label(self, canvas, label_number):
        """Tidy up the canvas after creating a label

        label_number - zero based number of the label just created

        """
        canvas.restoreState()

    def end_label_run(self, canvas, label_count):
        """Tidy up the canvas after a run of labels

        lable_count - number of labels in the run

        """
        canvas.showPage()


class LabelField(object):
    """Details of a single field on each label

    The value of the field is given as a format string so may contain boiler
    plate and zero or more attribues from the source table. The location and
    style of the field are also controlled.

    """

    X_COLUMN = "X"
    Y_COLUMN = "Y"
    WIDTH_COLUMN = "Width"
    HEIGHT_COLUMN = "Height"
    FONT_FAMILY_COLUMN = "Font family"
    FONT_SIZE_COLUMN = "Font size"
    COLOR_COLUMN = "Colour"
    LAYOUT_COLUMN = "Layout"
    FORMAT_COLUMN = "Format"

    COLUMNS = [
        X_COLUMN, Y_COLUMN, WIDTH_COLUMN, HEIGHT_COLUMN,
        FONT_FAMILY_COLUMN, FONT_SIZE_COLUMN,
        COLOR_COLUMN, LAYOUT_COLUMN, FORMAT_COLUMN,
    ]


    @classmethod
    def read_file(cls, csvfile, label_type):
        """Reads the field definition file and creates one instance per row

        csvfile - name of file to be read
        label_type - LabelType instance required to create LabelFields

        returns list of fields, these should be rendered in order that later
        fields can overlay earlier fields.

        """
        with open(csvfile, "r") as stream:
            reader = csv.reader(stream)
            header = next(reader)
            if cls.check_header(header) > 0:
                return []

            return [cls(dict(zip(header, row)), label_type)
                    for row in reader]

    @classmethod
    def check_header(cls, header):
        """Check tha the field file's header is valid

        """
        errors = 0
        for column_name in cls.COLUMNS:
            if column_name not in header:
                logger.error("Field file must have a '%s' column",
                             column_name)
                errors += 1
        return errors

    @classmethod
    def create_file(cls, csvfile):
        """Create an empty field definition file, just the headers are included

        """
        if os.path.exists(csvfile):
            logger.error("Field file '%s' exists, not overwriting",
                         csvfile)
            return

        with open(csvfile, 'w') as stream:
            writer = csv.writer(stream)
            writer.writerow(cls.COLUMNS)

    def __init__(self, parameters, label_type):
        """Create a LabelType instance for the given parameters

        """
        self.x = float(parameters[self.X_COLUMN]) * label_type.width
        self.y = float(parameters[self.Y_COLUMN]) * label_type.height
        self.width = float(parameters[self.WIDTH_COLUMN]) * label_type.width
        self.height = float(parameters[self.HEIGHT_COLUMN]) * label_type.height

        self.format = parameters[self.FORMAT_COLUMN]

        attributes = {}

        if parameters[self.COLOR_COLUMN]:
            attributes["textColor"] = colors.toColor(
                parameters[self.COLOR_COLUMN])

        if parameters[self.FONT_FAMILY_COLUMN]:
            attributes["fontName"] = parameters[self.FONT_FAMILY_COLUMN]

        if parameters[self.FONT_SIZE_COLUMN]:
            attributes["fontSize"] = (float(parameters[self.FONT_SIZE_COLUMN]) *
                                      label_type.height)

            attributes["leading"] = (float(parameters[self.FONT_SIZE_COLUMN]) *
                                     1.1 * label_type.height)

        if parameters[self.LAYOUT_COLUMN] == "c":
            attributes["alignment"] = styles.TA_CENTER

        self.style = styles.ParagraphStyle("", **attributes)

    def render(self, canvas, record):
        """Render this field on the canvas using data from the given record

        """
        text = self.format % record
        paragraph = Paragraph(text, self.style)

        frame = Frame(self.x, self.y, self.width, self.height,
                      topPadding=0, bottomPadding=0,
                      showBoundary=0)
        frame.addFromList([paragraph], canvas)


def read_csv(filename):
    """Generator to read a CSV file and yield name->value dictionaries

    One dictionary is yielded per row, each has the names taken from the header
    row of the CSV file.

    """
    with open(filename) as stream:
        reader = csv.reader(stream)
        header = next(reader)
        for row in reader:
            yield dict(zip(header, row))


def define_options(parser):
    """Add the command line option definitions to the parser

    """
    parser.add_option("--newlabelfile", action="store", dest="newlabelfile",
                      help="Create an empty CSV file for label types")
    parser.add_option("--labelfile", action="store", dest="label_file",
                      help="CSV file defining the label types")
    parser.add_option("--labeltype", action="store", dest="label_type",
                      help="Name of the label type to be used")

    parser.add_option("--newfieldfile", action="store", dest="newfieldfile",
                      help="Create an empty CSV file for label fields")
    parser.add_option("--fieldfile", action="store", dest="field_file",
                      help="CSV file defining the label fields")

    parser.add_option("--datafile", action="store", dest="data_file",
                      help="CSV file containing data for labels")

    parser.add_option("--outputfile", action="store", dest="output_file",
                      help="Name of PDF file to be created")

    parser.add_option("--quantity", action="store", dest="quantity",
                      help="Column holding quantity of given beer")
    parser.add_option("--enum", action="store", dest="enum",
                      help="Generated column for cask number")

    parser.set_defaults(label_file="labeltypes.csv",
                        field_file="labelfields.csv")


    parser.add_option("--debug", action="store_const", dest="loglevel",
                      const=logging.DEBUG,
                      help="Turn on debug logging")
    parser.add_option("--info", action="store_const", dest="loglevel",
                      const=logging.INFO,
                      help="Turn on info logging")
    parser.add_option("--warn", action="store_const", dest="loglevel",
                      const=logging.WARN,
                      help="Turn on warning logging")
    parser.add_option("--error", action="store_const", dest="loglevel",
                      const=logging.ERROR,
                      help="Turn on error logging")
    parser.add_option("--logformat", action="store", dest="logformat",
                      help="Format of log messages")
    parser.add_option("--logfile", action="store", dest="logfile",
                      help="Log file, stderr by default")
    parser.set_defaults(loglevel=logging.WARN, logformat=LOG_FORMAT)


def check_options(parser, options):
    """Check options for violations of requirement, mutually exclusion, etc.

    """
    if options.newlabelfile and (options.label_type or options.output_file):
        parser.error("--newlabelfile cannot be used with --labeltype or --outputfile")

    if options.newfieldfile and (options.label_type or options.output_file):
        parser.error("--newfieldfile cannot be used with --labeltype or --outputfile")


def process_cli():
    """Process the command line and return an Options instance

    """
    parser = optparse.OptionParser()
    define_options(parser)
    options, arguments = parser.parse_args()
    check_options(parser, options)
    return options


def setup_logging(options):
    """Perform logging basic config

    """
    logging.basicConfig(format=options.logformat,
                        level=options.loglevel,
                        filename=options.logfile)


def get_label_type(label_file, label_type_name):
    """Read the label type file and find the named type

    Logs error and exits if the required label type is not found

    label_file - path to the label type file
    label_type_name - name of the label type to be used

    returns LabelType instance

    """
    label_types = LabelType.read_file(label_file)

    if label_type_name not in label_types:
        logger.error("Label type '%s' not known", label_type_name)
        sys.exit(1)

    return label_types[label_type_name]


def add_labels(canvas, label_type, label_field_list, data_record_generator):
    """Add one label to the canvas for each data record

    canvas - Canvas instance, must be at the start of a new page
    label_type - LabelType instance
    label_field_list - list of LabelField instances
    data_record_generator - generator of name->value dictionaries

    """
    label_type.start_label_run(canvas)

    for i, data_fields in enumerate(data_record_generator):
        label_type.start_label(canvas, i)

        for label_field in label_field_list:
            label_field.render(canvas, data_fields)

        label_type.end_label(canvas, i)

    label_type.end_label_run(canvas, i + 1)


def create_labels(options):
    """Create the labels according to the options

    """
    label_type = get_label_type(options.label_file, options.label_type)
    label_fields = LabelField.read_file(options.field_file,
                                        label_type)

    canvas = Canvas(options.output_file)
    add_labels(canvas, label_type, label_fields, read_beers(options))
    canvas.save()

def read_beers(options):
    """Read beers from options.data_file, returns a generator

    If options.quantity and options.enum are set the column named by
    options.quantity should contain an integer indicating the number of casks
    of the beer.
    The generator will yield this number of entries with the column
    options.enum counting from 1.
    """
    generator = read_csv(options.data_file)

    if options.quantity and options.enum:
        return repeat_casks(options.quantity, options.enum, generator)
    return generator

def repeat_casks(quantity, enum, source):
    """Repeat each row the number of times incidated by the quantity column

    """
    for row in source:
        for cask in range(int(row[quantity])):
            row[enum] = cask + 1
            yield row

def main():
    """Create cask labels

    """
    options = process_cli()
    setup_logging(options)

    if options.newlabelfile or options.newfieldfile:
        if options.newlabelfile:
            LabelType.create_file(options.newlabelfile)
        if options.newfieldfile:
            LabelField.create_file(options.newfieldfile)
    else:
        create_labels(options)

if __name__ == "__main__":
    main()
