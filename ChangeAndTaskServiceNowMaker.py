# -*- coding: utf-8 -*-
import requests as rq
import tkinter as tk
from tkinter import filedialog
from requests.auth import HTTPBasicAuth
import xml.etree.ElementTree as ET
import xml.dom.minidom
import os
import sys
import time
import unicodedata
import glob
import csv
from datetime import timedelta, date
import json

# Parameters
url_sn_to_make_change = "https://synapsis.service-now.com/api/now/table/change_request"
url_sn_to_make_task = "https://synapsis.service-now.com/api/now/table/change_task"
url_sn_lists_changes = "https://synapsis.service-now.com/api/now/table/change_request?sysparm_query=active = true^stateNOT IN8, 3, 4^sys_created_byLIKErachid^ORsys_created_byLIKEbrayan"
url_sn_lists_task = "https://synapsis.service-now.com/api/now/table/change_task?sysparm_query=active = true ^ opened_byLIKERachid Moyse Polania ^ ORopened_byLIKEArthur Brayan Gallo Obando"
headers = {"Accept": "application/xml"}
user = 'rachid.moyse'
passwd = "Skills39**"
number = ""
root = tk.Tk("Archivo Nessus")
root.withdraw()
filepath = filedialog.askopenfilename()

csv.field_size_limit(sys.maxsize)


def NewTask(chgnumber, assigned_to, assigment_group, category, comments, description, short_description, reason, requested_by, state):
    payload_task = {
        'change_request': chgnumber,
        'assigned_to': assigned_to,
        'assignment_group': assigment_group,
        'category': category,
        'comments': comments,
        'description': description,
        'short_description': short_description,
        'reason': reason,
        'requested_by': requested_by,
        'u_fecha_incio_actividad': str(date.today()),
        'due_date': str(date.today() + timedelta(days=20)),
        'State': state,
    }
    headers_json = {"Content-Type": "application/json",
                    "Accept": "application/xml"}
    task_r = rq.post(url_sn_to_make_task, auth=HTTPBasicAuth(
        user, passwd), headers=headers_json, data=json.dumps(payload_task))

    file = open('data_temp_task.xml', 'w+')
    file.write(str(task_r.content)[2:-1])
    file.close()

    # Formatting
    dom4 = xml.dom.minidom.parse('data_temp_task.xml')
    pretty_xml_as_string4 = dom4.toprettyxml()
    file = open('data_temp_task.xml', 'w+')
    file.write(pretty_xml_as_string4)
    file.close()

    # Get Elements by Tagname
    tree4 = ET.parse('data_temp_task.xml')
    root4 = tree4.getroot()

    for task_tags in root4.findall('result'):
        number_task = task_tags.find('number').text
        sys_created_by_task = task_tags.find('sys_created_by').text
        assignment_group_task = task_tags.find(
            'assignment_group/link').text
        description_task = task_tags.find('description').text
        print("Numero de Tarea: ", number_task, "\nCreado por: ", sys_created_by_task,
              "\nAsignado al grupo: ", assignment_group_task, "\nDescripcion: ", description_task)
        print(
            "===================================================================================")


def NewChange(approval, risk, u_ambiente, assigned_to, assignment_group, category, description, short_description, impact, priority, reason, tpe):
    # Body to make a petition
    headers_xml = {"Content-Type": "application/json",
                   "Accept": "application/xml"}

    payload = {
        'approval': approval,
        'risk': risk,
        'u_ambiente': u_ambiente,
        'assigned_to': assigned_to,
        'assignment_group': assignment_group,
        'category': f'category',
        'comments': 'Solucion de vulnerabilidades',
        'description': description,
        'short_description': short_description,
        'impact': impact,
        'priority': priority,
        'reason': reason,
        'requested_by': user,
        'type': tpe,
    }

    post_r = rq.post(url_sn_to_make_change, auth=HTTPBasicAuth(
        user, passwd), headers=headers_xml, data=json.dumps(payload))

    file = open('data_temp.xml', 'w+')
    file.write(str(post_r.content)[2:-1])
    file.close()

    # Formatting
    dom3 = xml.dom.minidom.parse('data_temp.xml')
    pretty_xml_as_string3 = dom3.toprettyxml()
    file = open('data_temp.xml', 'w+')
    file.write(pretty_xml_as_string3)
    file.close()

    # Get Elements by Tagname
    tree3 = ET.parse('data_temp.xml')
    root3 = tree3.getroot()

    for number_change in root3.findall('result'):

        number = number_change.find('number').text
        sys_created_by = number_change.find('sys_created_by').text
        assignment_group = number_change.find('assignment_group/link').text
        description = number_change.find('description').text

        print("Numero de Cambio: ", number, "\nCreado por: ", sys_created_by,
              "\nAsignado al grupo: ", assignment_group, "\nDescripcion: ", description)
    return number


def CustomizeFile():
    reader = csv.DictReader(
        open(str(filepath)))
    with open(f"{os.getcwd()}\\CreacionCambiosyTareasfinal.csv", "w+") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=['Estado', 'Ambiente', 'Asignado A', 'Grupo Asignado', 'Categoria', 'Impacto', 'Prioridad',
                                                     'Motivo', 'SolicitadoPor', 'Tipo', 'Task-Categoria', 'Task-Comments', 'Task-State', 'Risk', 'Host', 'Protocol', 'Port', 'Name', 'Synopsis', 'Description', 'Solution', 'See Also'])
        writer.writeheader()
        for row in reader:
            writer.writerow({'Estado': '', 'Ambiente': '', 'Asignado A': '', 'Grupo Asignado': '', 'Categoria': '', 'Impacto': '', 'Prioridad': '',
                             'Motivo': '', 'SolicitadoPor': '', 'Tipo': '', 'Task-Categoria': '', 'Task-Comments': '', 'Task-State': '', 'Risk': row['Risk'], 'Host': row['Host'], 'Protocol': row['Protocol'], 'Port': row['Port'], 'Name': row['Name'], 'Synopsis': row['Synopsis'], 'Description': row['Description'], 'Solution': row['Solution'], 'See Also': row['See Also']})
        csvfile.close()
    print("[+]Se creo el archivo en el working directory!")


def PreviewFile():
    reader = csv.DictReader(
        open(f"{os.getcwd()}\\CreacionCambiosyTareasfinal.csv"))
    contador = 0
    current_host = 0
    for row in reader:
        next_host = row['Host']
        if current_host != next_host:
            print(
                "##########################################################################")
            change_number = f"Estado: {row['Estado']}\n Riesgo: {row['Risk']}\n Ambiente: {row['Ambiente']}\n Asignado A: {row['Asignado A']}\n Grupo Asignado: {row['Grupo Asignado']}\n Categoria: {row['Categoria']}\n Descripcion: {row['Description']}\n Short Descripcion: {row['Description']}\n Impacto: {row['Impacto']}\n Prioridad: {row['Prioridad']}\n Motivo: {row['Motivo']}\n Tipo: {row['Tipo']}"
            print(change_number)
            time.sleep(5)
            current_host = next_host
            contador = 0
        else:
            contador += 1
            print(f"Se crearan {contador} tarea(s) para el cambio de arriba")


def ExecuteFile():

    reader = csv.DictReader(
        open(str(glob.glob(f"{os.getcwd()}\\CreacionCambiosyTareasfinal.csv"))))
    current_host = 0
    for row in reader:
        next_host = row['Host']
        states = {"Critical": 1, "High": 2, "Medium": 3, "Low": 4}
        if current_host != next_host:
            change_number = NewChange(row['Estado'], states[row['Risk']], row['Ambiente'], row['Asignado A'], row['Grupo Asignado'],
                                      row['Categoria'], row['Description'], row['Description'], row['Impacto'], row['Prioridad'], row['Motivo'], row['Tipo'])
            time.sleep(10)
            current_host = next_host
        else:
            NewTask(change_number, row['Asignado A'],
                    row['Grupo Asignado'], row['Task-Categoria'], row['Task-Comments'], row['Description'], row['Description'], row['Motivo'], row['SolicitadoPor'], row['Task-State'])
            time.sleep(10)


CustomizeFile()
PreviewFile()
