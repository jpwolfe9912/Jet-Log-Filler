import os
import re

import fitz  # requires fitz, PyMuPDF
import pdfrw
import subprocess
import os.path
import sys
from PIL import Image

'''
    replace all the constants (the one in caps) with your own lists
'''

'''
FORM_KEYS is a dictionary (key-value pair) that contains 
1. keys - which are all the key names in the PDF form 
2. values - which are the type for all the keys in the PDF form. (string, checkbox, etc.)

Eg. PDF form contains 
1. First Name
2. Last Name
3. Sex (Male or Female)
4. Mobile Number

FORM_KEYS = {
    "fname": "string",
    "lname": "string",
    "sex": "checkbox",
    "mobile": "number"
}

This FORM_KEYS(key) returns the type of value for that key. 
I'm passing this as 2nd argument to encode_pdf_string() function.
'''
FORM_KEYS = {
    "other": "string",
    "route_to_1": "string",
    "route_to_2": "string",
    "route_to_3": "string",
    "route_to_4": "string",
    "route_to_5": "string",
    "route_to_6": "string",
    "route_to_7": "string",
    "route_to_8": "string",
    "route_to_9": "string",
    "route_to_10": "string",
    "route_to_11": "string",
    "route_to_alt_1": "string",
    "route_to_alt_2": "string",
    "route_to_alt_3": "string",
    "route_to_alt_4": "string",
    "route_to_alt_5": "string",
    "dep_aerodrome": "string",
    "dep_elev": "string",
    "dep_atis_id": "string",
    "dep_atis_freq": "string",
    "dest_aerodrome": "string",
    "dest_elev": "string",
    "alt_dest": "string",
    "alt_elev": "string",
    "chan_id_1": "string",
    "chan_freq_1": "string",
    "chan_id_2": "string",
    "chan_freq_2": "string",
    "chan_id_3": "string",
    "chan_freq_3": "string",
    "chan_id_4": "string",
    "chan_freq_4": "string",
    "chan_id_5": "string",
    "chan_freq_5": "string",
    "chan_id_6": "string",
    "chan_freq_6": "string",
    "chan_id_7": "string",
    "chan_freq_7": "string",
    "chan_id_8": "string",
    "chan_freq_8": "string",
    "chan_id_9": "string",
    "chan_freq_9": "string",
    "chan_id_10": "string",
    "chan_freq_10": "string",
    "chan_id_11": "string",
    "chan_freq_11": "string",
    "chan_id_alt_1": "string",
    "chan_freq_alt_1": "string",
    "chan_id_alt_2": "string",
    "chan_freq_alt_2": "string",
    "chan_id_alt_3": "string",
    "chan_freq_alt_3": "string",
    "chan_id_alt_4": "string",
    "chan_freq_alt_4": "string",
    "chan_id_alt_5": "string",
    "chan_freq_alt_5": "string",
    "course_1": "string",
    "course_2": "string",
    "course_3": "string",
    "course_4": "string",
    "course_5": "string",
    "course_6": "string",
    "course_7": "string",
    "course_8": "string",
    "course_9": "string",
    "course_10": "string",
    "course_11": "string",
    "course_alt_1": "string",
    "course_alt_2": "string",
    "course_alt_3": "string",
    "course_alt_4": "string",
    "course_alt_5": "string",
    "dep_clearance_id": "string",
    "dep_clearance_freq": "string",
    "time_off": "string",
    "dep_app_cont_id": "string",
    "dep_app_cont_freq": "string",
    "dist_1": "string",
    "dist_2": "string",
    "dist_3": "string",
    "dist_4": "string",
    "dist_5": "string",
    "dist_6": "string",
    "dist_7": "string",
    "dist_8": "string",
    "dist_9": "string",
    "dist_10": "string",
    "dist_11": "string",
    "dist_total": "string",
    "alt_route": "string",
    "alt_app_cont_id": "string",
    "alt_app_cont_freq": "string",
    "dist_alt_1": "string",
    "dist_alt_2": "string",
    "dist_alt_3": "string",
    "dist_alt_4": "string",
    "dist_alt_5": "string",
    "ete_1": "string",
    "ete_2": "string",
    "ete_3": "string",
    "ete_4": "string",
    "ete_5": "string",
    "ete_6": "string",
    "ete_7": "string",
    "ete_8": "string",
    "ete_9": "string",
    "ete_10": "string",
    "ete_11": "string",
    "ete_total": "string",
    "ete_alt_1": "string",
    "ete_alt_2": "string",
    "ete_alt_3": "string",
    "ete_alt_4": "string",
    "ete_alt_5": "string",
    "eta_1": "string",
    "ata_1": "string",
    "eta_2": "string",
    "ata_2": "string",
    "eta_3": "string",
    "ata_3": "string",
    "eta_4": "string",
    "ata_4": "string",
    "eta_5": "string",
    "ata_5": "string",
    "eta_6": "string",
    "ata_6": "string",
    "eta_7": "string",
    "ata_7": "string",
    "eta_8": "string",
    "ata_8": "string",
    "eta_9": "string",
    "ata_9": "string",
    "eta_10": "string",
    "ata_10": "string",
    "eta_11": "string",
    "ata_11": "string",
    "eta_total": "string",
    "ata_total": "string",
    "eta_alt_1": "string",
    "ata_alt_1": "string",
    "eta_alt_2": "string",
    "ata_alt_2": "string",
    "eta_alt_3": "string",
    "ata_alt_3": "string",
    "eta_alt_4": "string",
    "ata_alt_4": "string",
    "eta_alt_5": "string",
    "ata_alt_5": "string",
    "dep_gnd_cont_id": "string",
    "dep_gnd_cont_freq": "string",
    "tas": "string",
    "mach": "string",
    "dest_tower_id": "string",
    "dest_tower_freq": "string",
    "leg_fuel_1": "string",
    "leg_fuel_2": "string",
    "leg_fuel_3": "string",
    "leg_fuel_4": "string",
    "leg_fuel_5": "string",
    "leg_fuel_6": "string",
    "leg_fuel_7": "string",
    "leg_fuel_8": "string",
    "leg_fuel_9": "string",
    "leg_fuel_10": "string",
    "leg_fuel_11": "string",
    "leg_fuel_total": "string",
    "alt_altitude": "string",
    "alt_tower_id": "string",
    "alt_tower_freq": "string",
    "leg_fuel_alt_1": "string",
    "leg_fuel_alt_2": "string",
    "leg_fuel_alt_3": "string",
    "leg_fuel_alt_4": "string",
    "leg_fuel_alt_5": "string",
    "efr_1": "string",
    "afr_1": "string",
    "efr_2": "string",
    "afr_2": "string",
    "efr_3": "string",
    "afr_3": "string",
    "efr_4": "string",
    "afr_4": "string",
    "efr_5": "string",
    "afr_5": "string",
    "efr_6": "string",
    "afr_6": "string",
    "efr_7": "string",
    "afr_7": "string",
    "efr_8": "string",
    "afr_8": "string",
    "efr_9": "string",
    "afr_9": "string",
    "efr_10": "string",
    "afr_10": "string",
    "efr_11": "string",
    "afr_11": "string",
    "efr_total": "string",
    "afr_total": "string",
    "efr_alt_1": "string",
    "afr_alt_1": "string",
    "efr_alt_2": "string",
    "afr_alt_2": "string",
    "efr_alt_3": "string",
    "afr_alt_3": "string",
    "efr_alt_4": "string",
    "afr_alt_4": "string",
    "efr_alt_5": "string",
    "afr_alt_5": "string",
    "cont_fuel": "string",
    "cont_fuel_1": "string",
    "cont_fuel_2": "string",
    "cont_fuel_3": "string",
    "cont_fuel_4": "string",
    "cont_fuel_5": "string",
    "cont_fuel_6": "string",
    "cont_fuel_7": "string",
    "cont_fuel_8": "string",
    "cont_fuel_9": "string",
    "cont_fuel_10": "string",
    "cont_fuel_11": "string",
    "alt_fuel": "string",
    "cont_fuel_alt_1": "string",
    "cont_fuel_alt_2": "string",
    "cont_fuel_alt_3": "string",
    "cont_fuel_alt_4": "string",
    "cont_fuel_alt_5": "string",
    "dep_tower_id": "string",
    "dep_tower_freq": "string",
    "lbs_ph": "string",
    "lbs_pm": "string",
    "dest_gnd_cont_id": "string",
    "dest_gnd_cont_freq": "string",
    "notes_1": "string",
    "notes_2": "string",
    "notes_3": "string",
    "notes_4": "string",
    "notes_5": "string",
    "notes_6": "string",
    "notes_7": "string",
    "notes_8": "string",
    "notes_9": "string",
    "notes_10": "string",
    "notes_11": "string",
    "notes_12": "string",
    "alt_gnd_cont_id": "string",
    "alt_gnd_cont_freq": "string",
    "notes_alt_1": "string",
    "notes_alt_2": "string",
    "notes_alt_3": "string",
    "notes_alt_4": "string",
    "notes_alt_5": "string",
    "alt_time": "string",
    "route_dest_iaf_fuel": "string",
    "route_alt_iaf_fuel": "string",
    "approaches_fuel": "string",
    "in_air_used_fuel": "string",
    "reserve_fuel": "string",
    "rwy_length_dest": "string",
    "lighting_dest": "string",
    "fuel_dest": "string",
    "ils_dest": "string",
    "loc_dest": "string",
    "asr_dest": "string",
    "par_mins_dest": "string",
    "tac_mins_dest": "string",
    "arr_gear_dest": "string",
    "pubs_dest": "string",
    "notams_dest": "string",
    "fuel_packet_dest_1": "string",
    "fuel_packet_dest_2": "string",
    "fuel_packet_dest_3": "string",
    "fuel_packet_dest_4": "string",
    "etc_dest": "string",
    "last_cruise_req_fuel": "string",
    "map_to_iaf_req_fuel": "string",
    "bingo_req_fuel": "string",
    "last_cruise_appr_fuel": "string",
    "map_to_iaf_appr_fuel": "string",
    "rwy_length_alt": "string",
    "lighting_alt": "string",
    "fuel_alt": "string",
    "ils_alt": "string",
    "loc_alt": "string",
    "asr_alt": "string",
    "par_mins_alt": "string",
    "tac_mins_alt": "string",
    "arr_gear_alt": "string",
    "pubs_alt": "string",
    "notams_alt": "string",
    "fuel_packet_alt_1": "string",
    "fuel_packet_alt_2": "string",
    "fuel_packet_alt_3": "string",
    "fuel_packet_alt_4": "string",
    "etc_alt": "string",
    "last_cruise_res_fuel": "string",
    "map_to_iaf_fuel": "string",
    "add_res_fuel": "string",
    "stto_fuel": "string",
    "total_req_fuel": "string",
    "total_aboard_fuel": "string",
    "spare_fuel": "string",
    "last_cruise_total_fuel": "string",
    "map_to_iaf_total_fuel": "string",
    "bingo_total": "string",
    "waypoint_1": "string",
    "waypoint_2": "string",
    "waypoint_3": "string",
    "waypoint_4": "string",
    "waypoint_5": "string",
    "waypoint_6": "string",
    "waypoint_7": "string",
    "waypoint_8": "string",
    "waypoint_9": "string",
    "waypoint_10": "string",
    "waypoint_11": "string",
    "waypoint_12": "string",
    "waypoint_13": "string",
    "waypoint_14": "string",
    "waypoint_15": "string",
    "waypoint_16": "string",
    "clearance_cleared_to": "string",
    "clearance_altitude": "string",
    "clearance_freq": "string",
    "clearance_transp": "string",
    "clearance_route": "string"
}


def encode_pdf_string(value, type):
    if type == 'string':
        if value:
            return pdfrw.objects.pdfstring.PdfString.encode(value.upper())
        else:
            return pdfrw.objects.pdfstring.PdfString.encode('')
    elif type == 'checkbox':
        if value == 'True' or value == True:
            return pdfrw.objects.pdfname.BasePdfName('/Yes')
            # return pdfrw.objects.pdfstring.PdfString.encode('Y')
        else:
            return pdfrw.objects.pdfname.BasePdfName('/No')
            # return pdfrw.objects.pdfstring.PdfString.encode('')
    return ''


class ProcessPdf:

    def __init__(self, temp_directory, output_file):
        print('\n##########| Initiating Pdf Creation Process |#########\n')

        print('\nDirectory for storing all temporary files is: ', temp_directory)
        self.temp_directory = temp_directory
        print("Final Pdf name will be: ", output_file)
        self.output_file = output_file

    def add_data_to_pdf(self, template_path, data):
        print('\nAdding data to pdf...')
        template = pdfrw.PdfReader(template_path)

        for page in template.pages:
            annotations = page['/Annots']
            if annotations is None:
                continue

            for annotation in annotations:
                if annotation['/Subtype'] == '/Widget':
                    if annotation['/T']:
                        key = annotation['/T'][1:-1]
                        if re.search(r'.-[0-9]+', key):
                            key = key[:-2]

                        if key in data:
                            annotation.update(
                                pdfrw.PdfDict(V=encode_pdf_string(data[key], FORM_KEYS[key]))
                            )
                        annotation.update(pdfrw.PdfDict(Ff=1))

        template.Root.AcroForm.update(pdfrw.PdfDict(NeedAppearances=pdfrw.PdfObject('true')))
        pdfrw.PdfWriter().write(self.temp_directory + "data.pdf", template)
        print('Pdf saved')

        return self.temp_directory + "data.pdf"

    def convert_image_to_pdf(self, image_path, image_pdf_name):
        print('\nConverting image to pdf...')

        image = Image.open(image_path)
        image_rgb = image.convert('RGB')
        image_rgb.save(self.temp_directory + image_pdf_name)
        return self.temp_directory + image_pdf_name

    def add_image_to_pdf(self, pdf_path, images, positions):
        print('\nAdding images to Pdf...')

        file_handle = fitz.open(pdf_path)
        for position in positions:
            page = file_handle[int(position['page']) - 1]
            if not position['image'] in images:
                continue
            image = images[position['image']]
            page.insertImage(
                fitz.Rect(position['x0'], position['y0'], position['x1'], position['y1']),
                filename=image
            )

        file_handle.save(self.temp_directory + "data_image.pdf")
        print('images added')
        return self.temp_directory + "data_image.pdf"

    def delete_temp_files(self, pdf_list):
        print('\nDeleting Temporary Files...')
        for path in pdf_list:
            try:
                os.remove(path)
            except:
                pass

    def compress_pdf(self, input_file_path, power=3):
        """Function to compress PDF via Ghostscript command line interface"""
        quality = {
            0: '/default',
            1: '/prepress',
            2: '/printer',
            3: '/ebook',
            4: '/screen'
        }

        output_file_path = self.temp_directory + 'compressed.pdf'

        if not os.path.isfile(input_file_path):
            print("\nError: invalid path for input PDF file")
            sys.exit(1)

        if input_file_path.split('.')[-1].lower() != 'pdf':
            print("\nError: input file is not a PDF")
            sys.exit(1)

        print("\nCompressing PDF...")
        initial_size = os.path.getsize(input_file_path)
        subprocess.call(['gs', '-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4',
                         '-dPDFSETTINGS={}'.format(quality[power]),
                         '-dNOPAUSE', '-dQUIET', '-dBATCH',
                         '-sOutputFile={}'.format(output_file_path),
                         input_file_path]
                        )
        final_size = os.path.getsize(output_file_path)
        ratio = 1 - (final_size / initial_size)
        print("\nCompression by {0:.0%}.".format(ratio))
        print("Final file size is {0:.1f}MB".format(final_size / 1000000))
        return output_file_path
