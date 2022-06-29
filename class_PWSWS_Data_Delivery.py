import arcpy, os,zipfile,shutil

try:
        import win32com.client
except ImportError:
        arcpy.AddError("Unable to import python library win32com.client necessary for automated e-mail functions in Outlook")

class PWSWS_Data_Delivery:
        #mxd = arcpy.mapping.MapDocument(r'M:\PWSWS\XSHARE\Katie\LocatorMapTool\LocatorMap.mxd')
        #map_layers = arcpy.mapping.ListLayers(mxd)
        default_layers =["dbtgis_pwSWS.PWSWS.SAN_MAIN_CONNECTION_PT_ABAND",
                                         "dbtgis_pwSWS.PWSWS.SAN_LOCATE_CONNECTION_PT_ABAND",
                                         "dbtgis_pwSWS.PWSWS.SAN_MANHOLE_ABAND",
                                         "dbtgis_pwSWS.PWSWS.SAN_REGULATOR_ABAND",
                                         "dbtgis_pwSWS.PWSWS.SAN_MAIN_ABAND",
                                         "dbtgis_pwSWS.PWSWS.SAN_CONNECTION_ABAND",
                                         "dbtgis_pwSWS.PWSWS.SAN_BULKHEAD",
                                         "dbtgis_pwSWS.PWSWS.SAN_ENCASEMENT",
                                         "dbtgis_pwSWS.PWSWS.SAN_PILING",
                                         "dbtgis_pwSWS.PWSWS.SAN_PLANKING",
                                         "dbtgis_pwSWS.PWSWS.PACP_SAN_INSPECT_PT",
                                         "dbtgis_pwSWS.PWSWS.SAN_MISCELLANEOUS_LINE",
                                         "dbtgis_pwSWS.PWSWS.SAN_PROJECT_AREA",
                                         "dbtgis_pwSWS.PWSWS.SAN_MAIN_CONNECTION_PT",
                                         "dbtgis_pwSWS.PWSWS.SAN_CONNECTION",
                                         "dbtgis_pwSWS.PWSWS.SAN_LOCATE_CONNECTION_PT",
                                         "dbtgis_pwSWS.PWSWS.SAN_MANHOLE",
                                         "dbtgis_pwSWS.PWSWS.SAN_MAIN",
                                         "dbtgis_pwSWS.PWSWS.SAN_LIFT_STATION",
                                         "dbtgis_pwSWS.PWSWS.SAN_REGULATOR",
                                         "dbtgis_pwSWS.PWSWS.SAN_INFLOW_PT",
                                         "dbtgis_pwSWS.PWSWS.SAN_EVENT_PT",
                                         "dbtgis_pwSWS.PWSWS.SAN_GRADE_PT",
                                         "dbtgis_pwSWS.PWSWS.SAN_HYDRAULIC_BREAK",
                                         "dbtgis_pwSWS.PWSWS.SAN_DISCHARGE_PT",
                                         "dbtgis_pwSWS.PWSWS.SanitaryNetwork_Net_Junctions",
                                         "dbtgis_pwSWS.PWSWS.STORM_CB_RUN_ABAND",
                                         "dbtgis_pwSWS.PWSWS.STORM_MANHOLE_ABAND",
                                         "dbtgis_pwSWS.PWSWS.STORM_MAIN_ABAND",
                                         "dbtgis_pwSWS.PWSWS.STORM_MAIN_CONNECTION_PT_ABAND",
                                         "dbtgis_pwSWS.PWSWS.STORM_CONNECTION_ABAND",
                                         "dbtgis_pwSWS.PWSWS.STORM_LOCATE_CONNECTION_PT_ABAND",
                                         "dbtgis_pwSWS.PWSWS.STORM_OUTFALL_ABAND",
                                         "dbtgis_pwSWS.PWSWS.STORM_GRIT_CHAMBER_PRIVATE",
                                         "dbtgis_pwSWS.PWSWS.STORM_FILT_INFILT_DEVICE_PRIVATE",
                                         "dbtgis_pwSWS.PWSWS.STORM_STORAGE_STRUCTURE_PRIVATE",
                                         "dbtgis_pwSWS.PWSWS.STORM_RAIN_GAUGE",
                                         "dbtgis_pwSWS.PWSWS.STORM_MONITOR",
                                         "dbtgis_pwSWS.PWSWS.STORM_GRIT_CHAMBER_POLY",
                                         "dbtgis_pwSWS.PWSWS.STORM_FILT_INFILT_DEVICE_POLY",
                                         "dbtgis_pwSWS.PWSWS.STORM_STORAGE_STRUCTURE_POLY",
                                         "dbtgis_pwSWS.PWSWS.STORM_BULKHEAD",
                                         "dbtgis_pwSWS.PWSWS.PACP_STORM_INSPECT_PT",
                                         "dbtgis_pwSWS.PWSWS.STORM_ENCASEMENT",
                                         "dbtgis_pwSWS.PWSWS.STORM_PILING",
                                         "dbtgis_pwSWS.PWSWS.STORM_PLANKING",
                                         "dbtgis_pwSWS.PWSWS.STORM_MISCELLANEOUS_LINE",
                                         "dbtgis_pwSWS.PWSWS.STORM_PROJECT_AREA",
                                         "dbtgis_pwSWS.PWSWS.STORM_CB",
                                         "dbtgis_pwSWS.PWSWS.STORM_CB_RUN",
                                         "dbtgis_pwSWS.PWSWS.STORM_MANHOLE",
                                         "dbtgis_pwSWS.PWSWS.STORM_GRIT_CHAMBER",
                                         "dbtgis_pwSWS.PWSWS.STORM_EVENT_PT",
                                         "dbtgis_pwSWS.PWSWS.STORM_GRADE_PT",
                                         "dbtgis_pwSWS.PWSWS.STORM_HYDRAULIC_BREAK",
                                         "dbtgis_pwSWS.PWSWS.STORM_OUTFALL",
                                         "dbtgis_pwSWS.PWSWS.STORM_PUMP_STATION",
                                         "dbtgis_pwSWS.PWSWS.STORM_STORAGE_STRUCTURE",
                                         "dbtgis_pwSWS.PWSWS.STORM_MAIN",
                                         "dbtgis_pwSWS.PWSWS.STORM_MAIN_CONNECTION_PT",
                                         "dbtgis_pwSWS.PWSWS.STORM_CONNECTION",
                                         "dbtgis_pwSWS.PWSWS.STORM_LOCATE_CONNECTION_PT",
                                         "dbtgis_pwSWS.PWSWS.STORM_OPEN_CHANNEL",
                                         "dbtgis_pwSWS.PWSWS.STORM_OTHER_MS4_INFLOW_OUTFLOW",
                                         "dbtgis_pwSWS.PWSWS.STORM_PIPE_INLET_OUTLET",
                                         "dbtgis_pwSWS.PWSWS.SURFACE_WATER_POINT",
                                         "dbtgis_pwSWS.PWSWS.SWS_FLOWLINE",
                                         "dbtgis_pwSWS.PWSWS.StormwaterNetwork_Net_Junctions"]
        default_layers_split_names = [name.rsplit('.')[-1] for name in default_layers]
        determine_layers_selected = []
        column_extract_dictionary = {}
        final_records_copy = {}
        plats_copy_count = 0
        asbs_copy_count = 0
        owner_field_query = ['MPLS', 'MPRB', 'PRIVATE', 'None']
        outside_owners = []
        geodatabase_name = ""
        aoi_buffer_fc_path = ""
        try:
                o = win32com.client.Dispatch("Outlook.Application")
                namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
                accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts;
        except NameError:
                arcpy.AddError("Automated e-mail functions in Outlook unavailable")
        compressed_data = []
        email_outside_data = ""

        def zipper(dir, zip_file):
                zip = zipfile.ZipFile(zip_file, 'w', compression=zipfile.ZIP_DEFLATED,allowZip64=True)
                root_len = len(os.path.abspath(dir))
                for root, dirs, files in os.walk(dir):
                        archive_root = os.path.abspath(root)[root_len:]
                        for f in files:
                                fullpath = os.path.join(root, f)
                                archive_name = os.path.join(archive_root, f)
                                #print f
                                zip.write(fullpath, archive_name, zipfile.ZIP_DEFLATED)
                zip.close()
                return zip_file

        def __init__(self, output_path, data_type = None, input_selected_layers = None, use_input_selected_layers = False,recipient_name = None, recipient_email = None, sender_name = None, subject = None, e_mail_body = None):
                self.output_path = output_path
                if not os.path.exists(output_path):
                        os.makedirs(output_path)
                self.data_type = data_type
                self.input_selected_layers = input_selected_layers
                self.use_input_selected_layers = use_input_selected_layers
                self.recipient_name = recipient_name
                self.recipient_email = recipient_email
                self.sender_name = sender_name
                self.subject = subject
                self.e_mail_body = e_mail_body

        def create_aoi_as_feature_layer(self, polygon_coordinates = None, gsoc_ticket = None, company = None, worktype = None, contact = None, phone = None, work_for = None, email = None, caller = None, callerphone = None, sewer_laptop = None, sender = None, shapefiles = None):
                self.gsoc_ticket = gsoc_ticket
                self.contact = contact
                self.worktype = worktype
                self.work_for = work_for
                self.phone = phone
                self.company = company
                self.email = email
                self.caller = caller
                self.callerphone = callerphone
                self.sewer_laptop = sewer_laptop
                self.geodatabase_name = "data_request_polygon_" + self.gsoc_ticket + ".gdb"
                self.sender = sender
                self.shapefiles = shapefiles

                HEIpath = False

                # Create ticket gdb and set it as workspace.
                arcpy.CreateFileGDB_management(self.output_path, self.geodatabase_name)
                arcpy.env.workspace = os.path.join(self.output_path, self.geodatabase_name)
                # spatial reference code of NAD_1983_HARN_Adj_MN_Hennepin_Feet found here:
                # http://desktop.arcgis.com/en/arcmap/10.3/analyze/arcpy-classes/pdf/projected_coordinate_systems.pdf
                sr = arcpy.SpatialReference(104726)

                # Create ticket aoi layer
                aoi_feature_class = "aoi_polygon_" + gsoc_ticket
                arcpy.CreateFeatureclass_management(arcpy.env.workspace, aoi_feature_class, "POLYGON", spatial_reference = sr)
                with arcpy.da.InsertCursor(aoi_feature_class, ['SHAPE@']) as cursor:
                        cursor.insertRow([polygon_coordinates])

                # Create ticket aoi buffer layer
                aoi_buffer_fc = "aoi_polygon_buffer" + gsoc_ticket
                arcpy.Buffer_analysis(aoi_feature_class, aoi_buffer_fc, "15 Feet")
                self.aoi_buffer_fc_path = os.path.join(arcpy.env.workspace, aoi_buffer_fc)
                arcpy.AddMessage("\nAdding AOI polygon to map and selecting intersecting grid cells.")

                # Open title page mxd
                if HEIpath:
                        mxd_title = arcpy.mapping.MapDocument(r'D:\Development\MplsPythonScript\20211201\CodeNew\NonExcavationMapTool\TitlePage.mxd')
                else:
                        mxd_title = arcpy.mapping.MapDocument(r'G:\TOOLS\KatieLocatorTools\NonExcavationMapTool\TitlePage.mxd')
                df_title = arcpy.mapping.ListDataFrames(mxd_title, "La*")[0]

                # Add aoi buffer to title page toc and zoom to its extent
                aoi_feature_layer = self.aoi_buffer_fc_path
                select_feature_layer = arcpy.mapping.Layer(aoi_feature_layer)
                select_feature_layer.visible = True
                arcpy.mapping.AddLayer(df_title, select_feature_layer, "BOTTOM")
                lyr = arcpy.mapping.ListLayers(mxd_title, "aoi*")[0]
                ext = lyr.getExtent()
                df_title.extent = ext
                df_title.scale = df_title.scale * 1.3

                # Change title to show ticket number
                titleItem = arcpy.mapping.ListLayoutElements(mxd_title, "TEXT_ELEMENT", "title")[0]
                titleItem.text = ("GSOC Ticket " + self.gsoc_ticket)

                # Add ticket info text
                infoItem = arcpy.mapping.ListLayoutElements(mxd_title, "TEXT_ELEMENT", "TicketInfo")[0]
                infoItem.text = ("\nCaller: {}".format(self.caller) + ", {}".format(self.callerphone) + ", {}".format(self.email) + "\nCompany Info: {}".format(self.company) + "\nCompany Contact: {}".format(self.contact) + ", {}".format(self.phone) + "\nDone for: {}".format(self.work_for) + "\nType of work: {}".format(self.worktype))

                # save the title page pdf
                pdf_save = os.path.join(self.output_path, self.gsoc_ticket +"_map.pdf")
                map = pdf_save
                pdf = arcpy.mapping.PDFDocumentCreate(pdf_save)
                title = arcpy.mapping.ExportToPDF(mxd_title, map)
                pdf.appendPages(map)
                if HEIpath:
                        pdf.appendPages(r'D:\Development\MplsPythonScript\20211201\CodeNew\NonExcavationMapTool\legend.pdf')
                else:
                        pdf.appendPages('G:\TOOLS\KatieLocatorTools\LocatorMapTool\legend.pdf')
                arcpy.mapping.RemoveLayer(df_title, select_feature_layer)

                # create body pages of map grid cells
                # Open main map document and list data frames
                if HEIpath:
                        mxd2 = arcpy.mapping.MapDocument(r'D:\Development\MplsPythonScript\20211201\CodeNew\NonExcavationMapTool\nonexcavation.mxd')
                else:
                        mxd2 = arcpy.mapping.MapDocument(r'G:\TOOLS\KatieLocatorTools\NonExcavationMapTool\nonexcavation.mxd')
                df = arcpy.mapping.ListDataFrames(mxd2, "La*")[0]
                df2 = arcpy.mapping.ListDataFrames(mxd2, "In*")[0]

                # Add ticket aoi buffer layer to both dataframes.
                aoi_feature_layer = self.aoi_buffer_fc_path
                arcpy.AddMessage("{}".format(aoi_feature_layer))
                select_feature_layer = arcpy.mapping.Layer(aoi_feature_layer)
                select_feature_layer.visible = True
                arcpy.mapping.AddLayer(df, select_feature_layer, "BOTTOM")
                arcpy.mapping.AddLayer(df2, select_feature_layer, "BOTTOM")

                # Set title of map to ticket number
                ticketnum = arcpy.mapping.ListLayoutElements(mxd2, "TEXT_ELEMENT", "title")[0]
                ticketnum.text = ("GSOC Ticket: " + self.gsoc_ticket)

                # Create grid index feature layer
                ### Goes to the conditional so not making segments for single cell tickets
                grid_fc = "Segments_" + gsoc_ticket
                arcpy.GridIndexFeatures_cartography(grid_fc, select_feature_layer, 'INTERSECTFEATURE', 'NO_USEPAGEUNIT', '#', '514.648417 Feet', '658.455197 Feet')
                self.grid = os.path.join(arcpy.env.workspace, grid_fc)
                #df = arcpy.mapping.ListDataFrames(mxd, "La*")[0]
                grid_feature_layer = self.grid
                select_grid_layer = arcpy.mapping.Layer(grid_feature_layer)
                # Add layer to both data frames and set labels for the inset dataframe.
                ref = arcpy.mapping.ListLayers(mxd2, "Grid*", df)[0]
                arcpy.ApplySymbologyFromLayer_management(select_grid_layer, ref)
                select_grid_layer.visible = True
                ref.visible = False
                arcpy.mapping.AddLayer(df, select_grid_layer, "BOTTOM")
                arcpy.mapping.AddLayer(df2, select_grid_layer, "BOTTOM")
                grid_count1 = (arcpy.GetCount_management(select_grid_layer))
                grid_count = int(grid_count1.getOutput(0))
                arcpy.AddMessage("\n grid cells: {}".format(grid_count))

                if grid_count <= 1:
                        ticketnum.elementPositionX, ticketnum.elementPositionY = 4, 1.5
                        df2.elementPositionX, df2.elementPositionY = 11, 0
                        arcpy.mapping.AddLayer(df2, select_feature_layer, "BOTTOM")
                        arcpy.mapping.AddLayer(df, select_feature_layer, "BOTTOM")
                        arcpy.mapping.RemoveLayer(df, ref)
                        arcpy.mapping.RemoveLayer(df2, ref)
                        arcpy.AddMessage("Zooming to AOI")
##                        aoi_inset = arcpy.mapping.ListLayers(mxd2, "aoi*", df)[0]
##                        arcpy.AddMessage("layer to zoom {}".format(aoi_buffer_fc)) 
##                        aoi_extent = aoi_inset.getExtent()
##                        arcpy.AddMessage("extent {}".format(aoi_extent))
##                        df2.extent = aoi_extent
##                        df2.scale = df2.scale * 1.3
##                        arcpy.RefreshActiveView()
                        aoi_lyr = arcpy.mapping.ListLayers(mxd2, "aoi*", df2)[0]
                        #grid_lyr = arcpy.mapping.ListLayers(mxd2, "Grid*", df)[0]
                        ext = aoi_lyr.getExtent()
                        arcpy.AddMessage("extent {}".format(ext))
                        df.extent = ext
                        df.scale = df.scale * 1.2
                        arcpy.RefreshActiveView()
                        pdfName = os.path.join(self.output_path, gsoc_ticket + '.pdf')
                        arcpy.mapping.ExportToPDF(mxd2, pdfName)
                        pdf.appendPages(pdfName)
                        os.remove(pdfName)
                        arcpy.mapping.RemoveLayer(df, aoi_lyr)
                else:

                      grid_lyr = arcpy.mapping.ListLayers(mxd2, "Segments*", df2)[0]
                      if grid_lyr.supports("LABELCLASSES"):
                              for lblclass in grid_lyr.labelClasses:
                                      lblclass.showClassLabels = True
                                      lblclass.expression = '"{}" + [PageName] + "{}"'.format("<FNT size = '18'>","</FNT>")
                      grid_lyr.showLabels = True
        
                      # Zoom to the grid in the inset dataframe
                      ext2 = grid_lyr.getExtent()
                      df2.extent = ext2
                      df2.scale = df2.scale * 1
                      arcpy.RefreshActiveView()
                      # Make feature layer from grid index layer
                      grid_select = "Index_Grid_" + gsoc_ticket
                      arcpy.MakeFeatureLayer_management(select_grid_layer, grid_select)
        
                      # Add new grid feature layer to both data frames
                      selection = arcpy.mapping.Layer(grid_select)
                      arcpy.ApplySymbologyFromLayer_management(selection, ref)
                      arcpy.mapping.AddLayer(df, selection, "BOTTOM")
                      arcpy.mapping.AddLayer(df2, selection, "TOP")
                      arcpy.mapping.ListLayers(mxd2, df)
                      gridField = "PageName"
                      row, rows = None, None
                      rows = arcpy.SearchCursor(selection)
                      apnList = []
                      # List all rows in the grid index layer attribute table.
                      # Loop through each row and zoom to the specific pagename/cell area. export this extent as a single pdf map document. continue through all rows.
                      for row in rows:
                              apnList.append(row.getValue(gridField))
                      for APN in apnList:
                              whereClause = "PageName = '{0}'".format(APN)
                              arcpy.management.SelectLayerByAttribute(selection, "NEW_SELECTION",whereClause)
                              cellItem = arcpy.mapping.ListLayoutElements(mxd2, "TEXT_ELEMENT", "gridCell")[0]
                              cellItem.text = ("Current Grid Cell: {}".format(APN))
                              df.extent = selection.getSelectedExtent()
                              df.scale = df.scale * 1
                              arcpy.RefreshActiveView()
                              pdfName = os.path.join(self.output_path, gsoc_ticket + "_" + APN + '.pdf')
                              arcpy.mapping.ExportToPDF(mxd2, pdfName)
                              #pdf_mp = pdfName + '.pdf'
                              pdf.appendPages(pdfName)
                              os.remove(pdfName)
                              # Add disclaimer pdf to document
                if HEIpath:
                        pdf.appendPages(r'D:\Development\MplsPythonScript\20211201\CodeNew\NonExcavationMapTool\Sewer_Data_Disclaimer_6-28-17_(computer version)_No_Fees.pdf')
                else:
                        pdf.appendPages('G:\TOOLS\KatieLocatorTools\LocatorMapTool\Sewer_Data_Disclaimer_6-28-17_(computer version)_No_Fees.pdf')
                pdf.saveAndClose()

                #Make Shapefile
                if self.shapefiles:
                        shapefilePath = self.output_path + os.sep + 'shapefiles'
                        if not os.path.isdir(shapefilePath):
                                os.mkdir(shapefilePath)

                        featureclasses = arcpy.ListFeatureClasses()
                        arcpy.FeatureClassToShapefile_conversion(featureclasses, shapefilePath)

                        zip = zipfile.ZipFile(shapefilePath + '.zip', 'w', compression=zipfile.ZIP_DEFLATED,allowZip64=True)
                        root_len = len(os.path.abspath(shapefilePath))
                        for root, dirs, files in os.walk(shapefilePath):
                                archive_root = os.path.abspath(root)[root_len:]
                                for f in files:
                                        fullpath = os.path.join(root, f)
                                        archive_name = os.path.join(archive_root, f)
                                        #print f
                                        zip.write(fullpath, archive_name, zipfile.ZIP_DEFLATED)
                        zip.close()
                        shutil.rmtree(shapefilePath, ignore_errors=True)

                # create an email draft
                self.email = email
                obj = win32com.client.Dispatch("Outlook.Application")
                Msg = obj.CreateItem(0)
                Msg.To = self.email
                Msg.Body = "Dear " + self.caller + ",\n\nAttached is the data requested for the Area of Interest identified in the GSOC Ticket(s). The data includes a map of the areas sanitary and/or storm utilities in a pdf file. Also attached is a disclaimer for you to read to verify that you have received the data. Should you need any additional information, such as private laterals or GIS data, please do not hesitate to contact us again.\n\n Thank you, \n\n" + self.sender
                Msg.Subject = "City of Minneapolis Surface Water & Sewer Data Request #" + self.gsoc_ticket
                Msg.Attachments.Add(pdf_save)
                if self.shapefiles:
                        Msg.Attachments.Add(shapefilePath + '.zip')
                Msg.Save()
