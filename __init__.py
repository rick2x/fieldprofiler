# -*- coding: utf-8 -*-
"""
/***************************************************************************
 FieldProfilerPlugin
                                 A QGIS plugin
 Analyzes and summarizes attribute field values for selected layers.
 Generated by Plugin Builder: http://g-sherman.github.io/Qgis-Plugin-Builder/
                             -------------------
        begin                : 2025-05-08
        copyright            : (C) 2025 by ricks
        email                : rrcuario@gmail.com
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""

def classFactory(iface):
    """Load FieldProfilerPlugin class from file field_profiler_plugin.py.
    :param iface: A QGIS interface instance.
    :type iface: QgisInterface
    """
    from .field_profiler_plugin import FieldProfilerPlugin # Ensure this import is correct
    return FieldProfilerPlugin(iface)