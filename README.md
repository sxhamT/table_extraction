# table_extraction
Table extraction based on association with a specific heading


Libraries:
import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment
import glob
import numpy as np
from PyPDF2 import PdfReader, PdfWriter
import io
