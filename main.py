import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import sys
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn