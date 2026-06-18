from fastapi import FastAPI
import asyncio
import csv
from datetime import datetime
import json
import os
from typing import List, Optional, Union
import requests
from fastapi import Depends, FastAPI, HTTPException
from fastapi.security import OAuth2PasswordRequestForm
from numpy import conj
import pandas as pd
from pydantic import BaseModel, Field
import threading
import openpyxl
import numpy as np
import ast

# Initialize FastAPI app
app = FastAPI(
    title="My first API",
    description="API to just to say hello.",
    version="0.1.0",
)

@app.get("/hello", tags=["Greeting"])
def hello_fastapi():

    file_path="/workspaces/ngt039/ngt039/Worksheet in NGT_System_Design_v1.7.49.xlsx"

    # if os.path.exists(file_path):
    igw_df = pd.read_excel(file_path)  
    print(f"Loaded file: {file_path}")
    # else:
    #     igw_df = None  
    #     print(f"File {file_name} not found in {upload_folder_path}")

    refNo = "ngt039_501"
    taskRequester = "ngt039_501"
    taskRequest = "ngt039_501"
    oldNode = "ngt039_501"
    newNode = "ngt039_501"
    workplanFileName = "ngt039_501"
    workplanFileLink = "ngt039_501"
    taskStatus = "ngt039_501"
    serviceCode = "ngt039_501"
    condition_ip_transit = igw_df["service_list"].str.contains("IP-TRANSIT", na=False, case=False)
    condition_wia = igw_df["service_list"].str.contains("IP-TRANSIT", na=False, case=False) & \
                igw_df["service_list"].str.contains("WIA", na=False, case=False)

    # Filter rows where either condition is met
    igw_df = igw_df[condition_ip_transit | condition_wia]

    # Ensure igw_df contains required columns
    required_columns = ["routingType", "port", "vlan", "routeList", "bgp_neighbour_ipv4", "bgp_neighbour_ipv6"]
    if not all(col in igw_df.columns for col in required_columns):
        raise ValueError("Missing required columns in igw_df")

    empty_message = f"Service IP TRANSIT and WIA is not available in {oldNode}"

    # Prepare clean-up commands per row
    clean_up_data = []
    deactivate_data = []
    fallback_data = []

    if igw_df.empty:
        print("No valid data found in igw_df")
        clean_up_data = pd.DataFrame([[empty_message]])
        deactivate_data = pd.DataFrame([[empty_message]])
        fallback_data = pd.DataFrame([[empty_message]])

    else:   

        for _, row in igw_df.iterrows():
            routing_type = row["routingType"]
            port = row["port"]
            vlan = row["vlan"]
            route_list_raw = row["routeList"]
            bgp_ipv4 = row["bgp_neighbour_ipv4"]
            bgp_ipv6 = row["bgp_neighbour_ipv6"]
            description = row["interface_description"]
            service_list = row["service_list"]

            commands = []
            
            # Convert route_list from string to list if needed
            if isinstance(route_list_raw, str):
                try:
                    route_list = ast.literal_eval(route_list_raw)  # Convert string to list safely
                except (SyntaxError, ValueError):
                    route_list = [route_list_raw]  # Treat it as a single value if conversion fails
            else:
                route_list = route_list_raw if isinstance(route_list_raw, list) else [route_list_raw]

            # Ensure route_list is a valid list
            route_list = [route for route in route_list if isinstance(route, str) and route.strip()]

            if isinstance(description, list):
                description = description[0]
            new_description = description.strip("[]\"'") + f" migrated to {newNode}"


            """base_commands = [
                f"delete interfaces {port} unit {vlan}",
                f"delete class-of-service interfaces {port} unit {vlan}"
            ]

            base_deactivate_commands = [
                f"set interfaces {port} unit {vlan} description {new_description}",
                f"deactivate interfaces {port} unit {vlan}"
            ]

            base_fallback_commands = [
                f"activate interfaces {port} unit {vlan}",
                f"activate class-of-service interfaces {port} unit {vlan}"
            ]"""

            """if "IP-TRANSIT" in service_list and "WIA" not in service_list:

                if "Static" in routing_type:
                    static_commands_list = []
                    for route in route_list:
                        static_commands_list.append("\n".join(base_commands + [f"delete routing-options static route {route}"]))
                    clean_up_data.append({"Commands": "\n\n".join(static_commands_list)})  # Extra space between blocks

                    static_deactivate_list = []
                    for route in route_list:
                        static_deactivate_list.append("\n".join(base_deactivate_commands + [
                            f"deactivate routing-options static route {route}",
                            f"deactivate class-of-service interfaces {port} unit {vlan}"
                        ]))
                    deactivate_data.append({"Commands": "\n\n".join(static_deactivate_list)})  # Separate full sets by extra space

                    static_fallback_list = []
                    for route in route_list:
                        static_fallback_list.append("\n".join(base_fallback_commands + [
                            f"activate routing-options static route {route}"
                        ]))
                    fallback_data.append({"Commands": "\n\n".join(static_fallback_list)})

                elif "BGP" in routing_type:
                    bgp_commands = base_commands + [f"delete protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}"]
                    if pd.notna(bgp_ipv6) and bgp_ipv6:  # Only add IPv6 neighbor if it's not empty or null
                        bgp_commands.append(f"delete protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}")
                    clean_up_data.append({"Commands": "\n".join(bgp_commands)})

                    bgp_deactivate_commands = [
                        "deactivate interfaces",  # This line is separate from the port and vlan
                        f"deactivate protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}"
                    ]
                    if pd.notna(bgp_ipv6) and bgp_ipv6:
                        bgp_deactivate_commands.append(f"deactivate protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}")
                    bgp_deactivate_commands.append(f"{port} unit {vlan}")  # Move port & vlan to the last line
                    deactivate_data.append({"Commands": "\n".join(bgp_deactivate_commands)})

                    bgp_fallback_list = [
                        f"activate protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}"
                    ]
                    if pd.notna(bgp_ipv6) and bgp_ipv6:
                        bgp_fallback_list.append(f"activate protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}")
                    fallback_data.append({"Commands": "\n".join(base_fallback_commands + bgp_fallback_list)})

                elif "Direct" in routing_type:
                    clean_up_data.append({"Commands": "\n".join(base_commands)})

                    deactivate_data.append({"Commands": "\n".join(base_deactivate_commands + [
                        f"deactivate class-of-service interfaces {port} unit {vlan}"
                    ])})

                    fallback_data.append({"Commands": "\n".join(base_fallback_commands)})"""
            
            if "IP-TRANSIT" in service_list and "WIA" not in service_list:

                if "Static" in routing_type:
                    static_commands_list = []
                    for route in route_list:
                        commands = []
                        clean_up_data.extend([
                            f"delete interfaces {port} unit {vlan}",
                            f"delete routing-options static route {route}",
                            f"delete class-of-service interfaces {port} unit {vlan}",
                           ])
                        
                    #     commands.append(f"delete interfaces {port} unit {vlan}")
                    #     commands.append(f"delete routing-options static route {route}")
                    #     commands.append(f"delete class-of-service interfaces {port} unit {vlan}")
                    #     static_commands_list.append("\n".join(commands))
                    # clean_up_data.append({"Commands": "\n\n".join(static_commands_list)})

                    static_deactivate_list = []
                    for route in route_list:
                        commands = []
                        deactivate_data.extend([
                            f"set interfaces {port} unit {vlan} description {new_description}",
                            f"deactivate interfaces {port} unit {vlan}",
                            f"deactivate routing-options static route {route}",
                            f"deactivate class-of-service interfaces {port} unit {vlan}",
                           ])
                    #     commands.append(f"set interfaces {port} unit {vlan} description {new_description}")
                    #     commands.append(f"deactivate interfaces {port} unit {vlan}")
                    #     commands.append(f"deactivate routing-options static route {route}")
                    #     commands.append(f"deactivate class-of-service interfaces {port} unit {vlan}")
                    #     static_deactivate_list.append("\n".join(commands))
                    # deactivate_data.append({"Commands": "\n\n".join(static_deactivate_list)})

                    static_fallback_list = []
                    for route in route_list:
                        commands = []
                        fallback_data.extend([
                            f"activate interfaces {port} unit {vlan}",
                            f"activate routing-options static route {route}",
                            f"activate class-of-service interfaces {port} unit {vlan}",
                           ])
                    #     commands.append(f"activate interfaces {port} unit {vlan}")
                    #     commands.append(f"activate routing-options static route {route}")
                    #     commands.append(f"activate class-of-service interfaces {port} unit {vlan}")
                    #     static_fallback_list.append("\n".join(commands))
                    # fallback_data.append({"Commands": "\n\n".join(static_fallback_list)})

                elif "BGP" in routing_type:
                    commands = []
                    clean_up_data.extend([
                            f"delete interfaces {port} unit {vlan}",
                            f"delete protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}",
                        ])
                    
                    # commands.append(f"delete interfaces {port} unit {vlan}")
                    # commands.append(f"delete protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}")
                    if pd.notna(bgp_ipv6) and bgp_ipv6:
                        clean_up_data.extend([
                            f"delete protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}",
                        ])
                        # commands.append(f"delete protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}")
                    clean_up_data.extend([
                            f"delete class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"delete class-of-service interfaces {port} unit {vlan}")
                    # clean_up_data.append({"Commands": "\n".join(commands)})

                    commands = []
                    deactivate_data.extend([
                            f"deactivate interfaces {port} unit {vlan}",
                            f"deactivate protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}",
                        ])
                    # commands.append(f"deactivate interfaces {port} unit {vlan}")
                    # commands.append(f"deactivate protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}")
                    if pd.notna(bgp_ipv6) and bgp_ipv6:
                        deactivate_data.extend([
                            f"deactivate protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}",
                        ])
                        # commands.append(f"deactivate protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}")
                    deactivate_data.extend([
                            f"deactivate class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"deactivate class-of-service interfaces {port} unit {vlan}")
                    # deactivate_data.append({"Commands": "\n".join(commands)})

                    commands = []
                    fallback_data.extend([
                            f"activate interfaces {port} unit {vlan}",
                            f"activate protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}",
                        ])
                    # commands.append(f"activate interfaces {port} unit {vlan}")
                    # commands.append(f"activate protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}")
                    if pd.notna(bgp_ipv6) and bgp_ipv6:
                        fallback_data.extend([
                            f"activate protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}",
                        ])
                        # commands.append(f"activate protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}")
                    fallback_data.extend([
                            f"activate class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"activate class-of-service interfaces {port} unit {vlan}")
                    # fallback_data.append({"Commands": "\n".join(commands)})

                elif "Direct" in routing_type:
                    commands = []
                    clean_up_data.extend([
                            f"delete interfaces {port} unit {vlan}",
                            f"delete class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"delete interfaces {port} unit {vlan}")
                    # commands.append(f"delete class-of-service interfaces {port} unit {vlan}")
                    # clean_up_data.append({"Commands": "\n".join(commands)})

                    commands = []
                    deactivate_data.extend([
                            f"set interfaces {port} unit {vlan} description {new_description}",
                            f"deactivate interfaces {port} unit {vlan}",
                            f"deactivate class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"set interfaces {port} unit {vlan} description {new_description}")
                    # commands.append(f"deactivate interfaces {port} unit {vlan}")
                    # commands.append(f"deactivate class-of-service interfaces {port} unit {vlan}")
                    # deactivate_data.append({"Commands": "\n".join(commands)})

                    commands = []
                    fallback_data.extend([
                            f"activate interfaces {port} unit {vlan}",
                            f"activate class-of-service interfaces {port} unit {vlan}",
                        ])
                
                    # commands.append(f"activate interfaces {port} unit {vlan}")
                    # commands.append(f"activate class-of-service interfaces {port} unit {vlan}")
                    # fallback_data.append({"Commands": "\n".join(commands)})

            elif "WIA" in service_list:

                if "Static" in routing_type:
                    static_commands_list = []
                    for route in route_list:
                        commands = []
                        static_commands_list.extend([
                            f"delete interfaces {port} unit {vlan}",
                            f"delete routing-options static route {route}",
                            f"delete class-of-service interfaces {port} unit {vlan}",
                           ])
                        # commands.append(f"delete interfaces {port} unit {vlan}")
                        # commands.append(f"delete routing-options static route {route}")
                        # commands.append(f"delete class-of-service interfaces {port} unit {vlan}")
                        # static_commands_list.append("\n".join(commands))
                    # clean_up_data.append({"Commands": "\n\n".join(static_commands_list)})

                    static_deactivate_list = []
                    for route in route_list:
                        commands = []
                        static_deactivate_list.extend([
                            f"deactivate interfaces {port} unit {vlan}",
                            f"deactivate routing-options static route {route}",
                            f"deactivate class-of-service interfaces {port} unit {vlan}",
                           ])
                    #     commands.append(f"deactivate interfaces {port} unit {vlan}")
                    #     commands.append(f"deactivate routing-options static route {route}")
                    #     commands.append(f"deactivate class-of-service interfaces {port} unit {vlan}")
                    #     static_deactivate_list.append("\n".join(commands))
                    # deactivate_data.append({"Commands": "\n\n".join(static_deactivate_list)})

                    static_fallback_list = []
                    for route in route_list:
                        commands = []
                        static_fallback_list.extend([
                            f"activate interfaces {port} unit {vlan}",
                            f"activate routing-options static route {route}",
                            f"activate class-of-service interfaces {port} unit {vlan}",
                           ])
                    #     commands.append(f"activate interfaces {port} unit {vlan}")
                    #     commands.append(f"activate routing-options static route {route}")
                    #     commands.append(f"activate class-of-service interfaces {port} unit {vlan}")
                    #     static_fallback_list.append("\n".join(commands))
                    # fallback_data.append({"Commands": "\n\n".join(static_fallback_list)})

                elif "BGP" in routing_type:
                    commands = []
                    clean_up_data.extend([
                            f"delete interfaces {port} unit {vlan}",
                            f"delete protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}",
                        ])
                    # commands.append(f"delete interfaces {port} unit {vlan}")
                    # commands.append(f"delete protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}")
                    if bgp_ipv6:
                        clean_up_data.extend([
                            f"delete protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}",
                        ])
                        # commands.append(f"delete protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}")
                    clean_up_data.extend([
                            f"delete class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"delete class-of-service interfaces {port} unit {vlan}")
                    # clean_up_data.append({"Commands": "\n".join(commands)})

                    commands = []
                    deactivate_data.extend([
                            f"deactivate interfaces {port} unit {vlan}",
                            f"deactivate protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}",
                        ])
                    # commands.append(f"deactivate interfaces {port} unit {vlan}")
                    # commands.append(f"deactivate protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}")
                    if bgp_ipv6:
                        deactivate_data.extend([  
                            f"deactivate protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}",
                        ])
                        # commands.append(f"deactivate protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}")
                    deactivate_data.extend([
                            f"deactivate class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"deactivate class-of-service interfaces {port} unit {vlan}")
                    # deactivate_data.append({"Commands": "\n".join(commands)})

                    commands = []
                    fallback_data.extend([
                            f"activate interfaces {port} unit {vlan}",
                            f"activate protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}",
                        ])
                    # commands.append(f"activate interfaces {port} unit {vlan}")
                    # commands.append(f"activate protocols bgp group eBGP_IPv4 neighbor {bgp_ipv4}")
                    if bgp_ipv6:
                        fallback_data.extend([
                            f"activate protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}",
                        ])
                        # commands.append(f"activate protocols bgp group eBGP_IPv6 neighbor {bgp_ipv6}")
                    fallback_data.extend([
                            f"activate class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"activate class-of-service interfaces {port} unit {vlan}")
                    # fallback_data.append({"Commands": "\n".join(commands)})

                elif "Direct" in routing_type:
                    commands = []
                    clean_up_data.extend([
                            f"delete interfaces {port} unit {vlan}",
                            f"delete class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"delete interfaces {port} unit {vlan}")
                    # commands.append(f"delete class-of-service interfaces {port} unit {vlan}")
                    # clean_up_data.append({"Commands": "\n".join(commands)})

                    commands = []
                    deactivate_data.extend([
                            f"deactivate interfaces {port} unit {vlan}",
                            f"deactivate class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"deactivate interfaces {port} unit {vlan}")
                    # commands.append(f"deactivate class-of-service interfaces {port} unit {vlan}")
                    # deactivate_data.append({"Commands": "\n".join(commands)})

                    commands = []
                    fallback_data.extend([
                            f"activate interfaces {port} unit {vlan}",
                            f"activate class-of-service interfaces {port} unit {vlan}",
                        ])
                    # commands.append(f"activate interfaces {port} unit {vlan}")
                    # commands.append(f"activate class-of-service interfaces {port} unit {vlan}")
                    # fallback_data.append({"Commands": "\n".join(commands)})

    # Convert to DataFrame
    clean_up_df = pd.DataFrame(clean_up_data)
    deactivate_df = pd.DataFrame(deactivate_data)
    fallback_df = pd.DataFrame(fallback_data)

    OUTPUT_FILE = "/workspaces/ngt039/ngt039/output_039.xlsx"
    # Save to Excel
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        clean_up_df.to_excel(writer, sheet_name="clean up", index=False, header=False)
        deactivate_df.to_excel(writer, sheet_name="deactivate", index=False, header=False)
        fallback_df.to_excel(writer, sheet_name="fallback", index=False, header=False)  # Empty sheet

        
    return {"message": "Hello, NGT_039!"}