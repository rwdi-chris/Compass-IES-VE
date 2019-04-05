"""compass_export.py: Export script to be used in IES-VE VEScript environment to export data for use on EnergyCompass.design"""
__author__ = 'Chris Frankowski'
__email__ = 'chris.frankowski@rwdi.com'
__version__ = '2018.0.0'

import iesve, json, os
import numpy as np
from ies_file_picker import IesFilePicker
from tkinter import Tk, simpledialog, messagebox, filedialog

PROGRESS_BAR_WIDTH = 100
HDD_REF = 18.3


def export():
    proposed_inputs = get_user_inputs('Proposed')
    proposed_results = get_results(proposed_inputs, 'Proposed')

    export_data = {
        'exporter_version' : __version__,
        'proposed_results' : proposed_results,
    }

    root = Tk()
    root.withdraw()
    root.lift()
    root.focus_force()

    attach_reference = messagebox.askyesno("Attach Reference Model?", '{:^100}'.format("Attach Reference Model?"))
    root.destroy()

    if attach_reference:
        reference_inputs = get_user_inputs('Reference')
        reference_inputs['orientation'] = proposed_inputs['orientation']
        reference_results = get_results(reference_inputs, 'Reference')
        export_data['reference_results'] = reference_results

    suffix = ' [both]' if attach_reference else ' [proposed]'

    write_file(export_data, suffix)


def get_user_inputs(model_type):
    root = Tk()
    root.withdraw()
    root.overrideredirect(True)
    root.geometry('0x0+0+0')
    root.deiconify()
    root.lift()
    root.focus_force()

    results = {}

    if model_type == 'Proposed':
        results['orientation'] = simpledialog.askinteger("Orientation", '{:^100}'.format("Enter the model orientation"),
                                                         parent=root, minvalue=0, maxvalue=360)

    room_nodes = simpledialog.askstring("Room Nodes",
                                        '{:^100}'.format("(" + model_type + " Model) Enter Room Supply Air Nodes"),
                                        parent=root)
    if room_nodes:
        results['room_nodes'] = [int(node) for node in room_nodes.split(',')]
    else:
        results['room_nodes'] = []

    oa_intake_nodes = simpledialog.askstring("Outside Air Intake Nodes",
                                             '{:^100}'.format("(" + model_type + " Model) Enter Outside Air Intake Nodes"),
                                             parent=root)
    if oa_intake_nodes:
        results['oa_intake_nodes'] = [int(node) for node in oa_intake_nodes.split(',')]
    else:
        results['oa_intake_nodes'] = []

    results['file_name'] = IesFilePicker.pick_aps_file()

    root.destroy()

    return results


def get_results(user_input, model_type):
    results = {}
    file_name = user_input['file_name']

    with iesve.ResultsReader.open(file_name) as aps_file:
        results['energy_uses'] = get_energy(aps_file)
        results['aps_stats'] = get_aps_stats(aps_file)
        results['weather'] = get_weather(aps_file)
        results['location'] = get_location()
        results['location']['orientation'] = user_input['orientation']
        results['building_results'] = get_building_results(aps_file)
        results['gains'] = get_gains(model_type)
        results['bodies'] = get_bodies(model_type)
        results['costs'] = get_costs(file_name)
        results['airflows'] = get_airflows(aps_file, user_input['room_nodes'], user_input['oa_intake_nodes'])
        results['room_results'] = get_room_results(aps_file)
        results['diagnostic'] = {'room_nodes': user_input['room_nodes'], 'oa_intake_nodes': user_input['oa_intake_nodes']}

    return results


def get_energy(aps_file):
    print('Gathering Energy Uses...', end='')
    energy_uses = aps_file.get_energy_uses()
    energy_sources = aps_file.get_energy_sources()
    energy_uses_export = {}
    for use_k, use_v in energy_uses.items():
        energy_sources_export = {}
        for source_k, source_v in energy_sources.items():
            result = aps_file.get_energy_results(use_id=use_k, source_id=source_k)
            energy_usage = np.sum(result)
            if energy_usage:
                energy_usage = energy_usage * (24 / aps_file.results_per_day) / 1000
                demand = np.max(result)
                energy_sources_export[str(source_k)] = {'name': source_v['name'], 'cef': source_v['cef'],
                                                        'usage': energy_usage, 'demand': demand / 1000,
                                                        'all': [round(float(r), 2) for r in result]}
        energy_uses_export[str(use_k)] = {'name': use_v['name'], 'sources': energy_sources_export}
    print('DONE')
    return energy_uses_export


def get_aps_stats(aps_file):
    print('Gathering Simulation Stats...', end='')
    sizes = aps_file.get_conditioned_sizes()
    aps_stats = {
        'first_day' : aps_file.first_day,
        'last_day' : aps_file.last_day,
        'results_per_day' : aps_file.results_per_day,
        'weather_file' : aps_file.weather_file,
        'year' : aps_file.year,
        'sizes' : {
            'area': sizes[0],
            'volume': sizes[1],
            'rooms': sizes[2]
        },
        'VE_version' : iesve.VEProject.get_current_project().get_version(),
    }

    print('DONE')
    return aps_stats


def get_weather(aps_file):
    print('Gathering Weather Data...', end='')
    temps = aps_file.get_weather_results('Temperature', 'Dry-bulb temperature')

    hdd = sum([max(HDD_REF - temp, 0) for temp in temps]) / aps_file.results_per_day
    cdd = sum([max(temp - HDD_REF, 0) for temp in temps]) / aps_file.results_per_day
    print('DONE')
    return {'heating_degree_days': hdd, 'cooling_degree_days': cdd}


def get_location():
    print('Gathering Location Data...', end='')
    loc = iesve.VELocate()
    loc.open_wea_data()
    data = loc.get()
    loc.close_wea_data()
    print('DONE')
    return data


def get_bodies(model_type):
    ve_project = iesve.VEProject.get_current_project()
    if model_type == 'Proposed':
        model = ve_project.models[0]
    else:
        model = ve_project.models[1]

    bodies = [body for body in model.get_bodies(False) if body.type == iesve.VEBody_type.room]

    constructions_set = set()
    bodies_output = []

    print("\nScanning constructions")
    pb_header()
    for body_i, body in enumerate(bodies):
        body_constructions = [c[0] for c in body.get_assigned_constructions()]
        body_areas = body.get_areas()
        body_areas_key_map = [
            ('int_floor_area','ifa'),('int_floor_glazed','ifg'),('int_floor_opening','ifo'),
            ('int_ceiling_area','ica'),('int_ceiling_glazed','icg'),('int_ceiling_opening','ico'),('int_ceiling_door','icd'),
            ('int_wall_area','iwa'),('int_wall_glazed','iwg'),('int_wall_opening','iwo'),('int_wall_door','iwd'),
            ('ext_floor_area','efa'),('ext_floor_glazed','efg'),('ext_floor_opening','efo'),
            ('ext_ceiling_area','eca'),('ext_ceiling_glazed','ecg'),('ext_ceiling_opening','eco'),('ext_ceiling_door','ecd'),
            ('ext_wall_area','ewa'),('ext_wall_glazed','ewg'),('ext_wall_opening','ewo'),('ext_wall_door','ewd'),
            ('volume','v')
        ]
        body_areas = reduce_dict(body_areas,body_areas_key_map)
        
        surfaces = body.get_surfaces()
        surface_output = []
        
        pb_update(body_i, len(bodies))
        for sur in surfaces:
            adjacencies = sur.get_adjacencies()
            adjacencies_key_map = [('gross', 'g'), ('hole', 'h'), ('door', 'd'), ('window', 'w')]
            adjacency_output = []
            for adjacency in adjacencies:
                adjacency_properties = adjacency.get_properties()
                adjacency_properties = reduce_dict(adjacency_properties, adjacencies_key_map)
                adjacency_construction = adjacency.get_construction()
                adjacency_properties['c'] = adjacency_construction
                constructions_set.add(adjacency_construction)
                adjacency_output.append(adjacency_properties)

            areas = sur.get_areas()
            areas_key_map = [('total_gross', 'tg'), ('total_net', 'tn'), ('total_window', 'tw'),
                             ('total_door', 'td'), ('total_hole', 'th'), ('total_gross_openings', 'tgo'),
                             ('internal_gross', 'ig'), ('internal_net', 'in'), ('internal_window', 'iw'),
                             ('internal_door', 'id'), ('internal_hole', 'ih'), ('internal_gross_openings', 'igo'),
                             ('external_gross', 'eg'), ('external_net', 'en'), ('external_window', 'ew'),
                             ('external_door', 'ed'), ('external_hole', 'eh'), ('external_gross_openings', 'ego')]
            areas = reduce_dict(areas, areas_key_map)

            constructions = sur.get_constructions()
            for c in constructions:
                constructions_set.add(c)

            opening_totals = sur.get_opening_totals()
            opening_totals_key_map = [('openings', 'o'), ('holes', 'h'), ('doors', 'd'), ('windows', 'w'),
                                      ('external_holes', 'eh'), ('external_doors', 'ed'), ('external_windows', 'ew')]
            opening_totals = reduce_dict(opening_totals, opening_totals_key_map)

            properties = sur.get_properties()
            properties_key_map = [('type', 'ty'), ('area', 'a'), ('orientation', 'o'), ('tilt', 'ti')]
            properties = reduce_dict(properties, properties_key_map)

            surface_output.append({
                'adjacencies': adjacency_output,
                'areas': areas,
                'openings': opening_totals,
                'properties': properties,
                'constructions' : constructions,
            })
        
        bodies_output.append({
        'id' : body.id,
        'construction': body_constructions,
        'surfaces' : surface_output,
        'areas' : body_areas,
        'subtype': str(body.subtype),
        })

    # get the Project (type=0) tuple (this is what we are normally interested in, the project list associated with the VE model)
    # this tuple will always have a project list of length 1, the only project associated with the VE model
    constructions_output = {}
    constructions_db = iesve.VECdbDatabase.get_current_database().get_projects()[0][0]

    for c in constructions_set:
        construction = constructions_db.get_construction(c, iesve.construction_class.none)
        u_value = construction.get_u_factor(iesve.uvalue_types.ashrae)

        constructions_output[construction.id] = {
            'category': str(construction.category),
            'u_value': round(u_value, 6),
            'reference' : construction.reference,
        }

    return {'constructions': constructions_output, 'bodies': bodies_output}


def get_room_results(aps_file):
    room_list = aps_file.get_room_list()
    result_length = ((aps_file.last_day - aps_file.first_day + 1) * aps_file.results_per_day)

    heat_excluding_oa = np.zeros(result_length)
    cool_excluding_oa = np.zeros(result_length)

    print("Gathering Room Results")
    pb_header()
    for room_i, room in enumerate(room_list):
        pb_update(room_i, len(room_list))

        internal_gain = aps_file.get_room_results(room[1], 'Casual gains', 'Internal gain',
                                                  'z') + aps_file.get_room_results(room[1], 'Internal latent gain',
                                                                                   'Internal latent gain', 'z')
        solar_gain = aps_file.get_room_results(room[1], 'Window solar gains', 'Solar gain', 'z')
        infiltration_gain = aps_file.get_room_results(room[1], 'Infiltration gain', 'Infiltration gain', 'z')
        infiltration_gain_lat = aps_file.get_room_results(room[1], 'Infiltration lat gain', 'Infiltration lat gain', 'z')
        if infiltration_gain_lat:
            infiltration_gain += infiltration_gain_lat
        external_conduction_gain = aps_file.get_room_results(room[1], 'Conduction from ext elements',
                                                             'External conduction gain', 'z')
        internal_conduction_gain = aps_file.get_room_results(room[1], 'Conduction from int surfaces',
                                                             'Internal conduction gain', 'z')

        total_gain = internal_gain + solar_gain + infiltration_gain + external_conduction_gain + internal_conduction_gain

        for i, v in enumerate(total_gain):
            if v > 0:
                cool_excluding_oa[i] += -v / 1000
            elif v < 0:
                heat_excluding_oa[i] += -v / 1000

    return {'heat_excluding_oa': max(heat_excluding_oa), 'cool_excluding_oa': min(cool_excluding_oa)}


def get_building_results(aps_file):
    print("Gathering Building Results...")
    variables = aps_file.get_variables()
    results = {}

    pb_header()
    for i, var in enumerate(variables):
        pb_update(i, len(variables))
        if var['units_type'] in ['Power', 'Sys Load']:

            total = np.sum(aps_file.get_results(var['aps_varname'], var['display_name'], var['model_level']))
            peak = np.max(aps_file.get_results(var['aps_varname'], var['display_name'], var['model_level']))

            if total:
                results[var['aps_varname']] = {'total': float(total), 'peak': float(peak)}

    return results


def get_costs(aps_file):
    str_info_message = iesve.TariffsEngine.String()
    str_error = iesve.TariffsEngine.String()
    tariff_engine = iesve.TariffsEngine()

    aps_design_file_path = aps_file
    aps_benchmark_file_path = ''

    tariff_engine.Init(iesve.VEProject.get_current_project().path,
                aps_design_file_path,
                aps_benchmark_file_path,
                str_info_message,
                str_error,
                iesve.TariffsEngine.EUnitsSystem.METRIC,
                iesve.TariffsEngine.EModes.MODE_NORMAL,
                iesve.TariffsEngine.EEnergyDataset.ENERGY_DATASET_ASHRAE,
                iesve.TariffsEngine.EComputeCosts.COMPUTE_COSTS_YES)

    # Check for errors/warnings in initialisation
    if not str_error.Empty():
        print("Error:", str_error.GetString())
    if not str_info_message.Empty():
        print("Info:", str_info_message.GetString())  # print info but continue anyway

    utilities = tariff_engine.GetUtilitiesNamesAndIds()
    output = {}

    for utility in utilities:
        output[utility[0]] = tariff_engine.GetDesignNetCost(utility[1])

    return output


def get_gains(model_type):
    print("Gathering Gain Data...", end="")
    ve_project = iesve.VEProject.get_current_project()
    if model_type == 'Proposed':
        model = ve_project.models[0]
        get_room_data_type = 0
    else:
        model = ve_project.models[1]
        get_room_data_type = 2

    bodies = model.get_bodies(False)  # [:]

    room_gains = {'Lighting': 0, 'People': 0, 'Equipment': 0}

    for body in bodies:
        gains = body.get_room_data(get_room_data_type).get_internal_gains()  # this needs to be fixed
        for gain in gains:
            results = gain.get()
            if results['type_str'] in ['Machinery', 'Miscellaneous', 'Cooking', 'Computers']:
                room_gains['Equipment'] += results['max_power_consumptions'][1]
            elif results['type_str'] in ['Fluorescent Lighting', 'Tungsten Lighting']:
                room_gains['Lighting'] += results['max_power_consumptions'][1]
            elif results['type_str'] in ['People']:
                room_gains['People'] += results['occupancies'][1]
    print("DONE")
    return room_gains


def get_airflows(aps_file, room_nodes, oa_intake_nodes):
    room_results = []

    print("Gathering Room airflows")
    pb_header()
    for i, room in enumerate(room_nodes):
        pb_update(i, len(room_nodes))
        room_results.append(ghnr(aps_file, room))

    oa_results = []
    print("Gathering Outside Air airflows")
    pb_header()
    for i, oa_intake in enumerate(oa_intake_nodes):
        pb_update(i, len(oa_intake_nodes))
        oa_results.append(ghnr(aps_file, oa_intake))

    room_grand_total = 0
    room_hourly_total = None
    oa_grand_total = 0
    oa_hourly_total = None

    for room in room_results:
        if room is not None:
            if room_hourly_total is None:
                room_hourly_total = room
            else:
                room_hourly_total += room
            room_grand_total += sum(room) * (3600 * (24 / aps_file.results_per_day))

    for oa in oa_results:
        if oa is not None:
            if oa_hourly_total is None:
                oa_hourly_total = oa
            else:
                oa_hourly_total += oa
            oa_grand_total += sum(oa) * (3600 * (24 / aps_file.results_per_day))

    if room_hourly_total is not None and oa_hourly_total is not None:
        return {'supply_air_total': float(room_grand_total), 'supply_air_max': float(max(room_hourly_total)),
                'outside_air_total': float(oa_grand_total), 'outside_air_max': float(max(oa_hourly_total))}
    else:
        return {'supply_air_total': float(room_grand_total), 'supply_air_max': 0,
                'outside_air_total': float(oa_grand_total), 'outside_air_max': 0}


def ghnr(aps_file, node):
    if aps_file.get_hvac_node_results(node, -1, 'Volume flow') is None and aps_file.get_hvac_node_results(node, 1,
                                                                                                          'Volume flow') is None:
        return None  # Node doesn't exist
    elif aps_file.get_hvac_node_results(node, -1, 'Volume flow') is not None:
        return aps_file.get_hvac_node_results(node, -1, 'Volume flow')  # Node is a plant side node
    else:
        layer = 1
        total = aps_file.get_hvac_node_results(node, 1, 'Volume flow')
        while True:
            layer += 1
            result = aps_file.get_hvac_node_results(node, layer, 'Volume flow')
            if result is None:
                break
            else:
                total += result
        return total


def pb_header():
    global _threshold
    _threshold = 0.01
    print("[" + "progress".center(100) + "]\n[", end="")


def pb_update(index, length):
    global _threshold
    while round(((index + 1) / length), 3) >= round(_threshold, 3):
        print("X", end="")
        _threshold += 0.01
    if index + 1 == length:
        print("]")


def reduce_dict(full_dict, key_map):
    new_dict = {}
    for lk, sk in key_map:
        if full_dict[lk] != 0:
            try:
                new_dict[sk] = round(full_dict[lk], 2)
            except TypeError:
                new_dict[sk] = full_dict[lk]
    return new_dict


def write_file(export_data, suffix):
    root = Tk()
    root.withdraw()
    file_path = None
    project_name = iesve.VEProject.get_current_project().name

    while not file_path:
        file_path = filedialog.asksaveasfilename(initialfile='EC.d Export ' + project_name + suffix + '.json',
                                                 title="Save Results as", filetypes=(("JSON", "*.json"),))
    if os.path.splitext(file_path)[1] != '.json':
        file_path += '.json'

    root.destroy()

    with open(file_path, 'w') as export_file:
        json.dump(export_data, export_file)
    print("Export created: " + file_path)


if __name__ == '__main__':
    if iesve.VEProject.get_current_project().name == 'Untitled':
        print("Please open the project before running the export script.")
    else:
        export()
