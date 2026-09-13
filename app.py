import streamlit as st
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import json, os, base64, zipfile
from datetime import date, datetime
from pypdf import PdfReader, PdfWriter
from PIL import Image
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from batch_generate import render_batch_tab

TEMPLATE_UNITED = "backflow_template.pdf"
TEMPLATE_JAX = "jacksonville_template.pdf"
TECHNICIANS_FILE = "technicians.json"
PAGE_W, PAGE_H = 612, 792
JAX_PAGE_W, JAX_PAGE_H = 612, 792