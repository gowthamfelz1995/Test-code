# importing model classes
from datetime import date
from Models.objwrapper import obj_wrap
from Models.fieldwrap import field_wrap_obj
from Models.grandwrap import grand_wrap_obj
from Models.parentwrap import parent_wrap_obj
from Models.childwrap import child_wrap_obj

import re
import json

# The locale module opens access to the POSIX locale database and functionality
import locale
# user’s preferred locale settings
locale.setlocale(locale.LC_ALL, 'en_US.utf-8')


class utility:
    # Method to return index of the obj
    def check_obj_present(self, current_obj, child_wrapper):
        obj_list = list()
        for record in child_wrapper.parentObjWrapperList:
            obj_list.append(record.objName)
        if len(obj_list) > 0:
            try:
                return obj_list.index(current_obj)
            except ValueError:
                return -1
        else:
            return -1

    # Method to return index of the parent obj
    def check_grand_obj_present(self, current_obj, child_wrapper, parent_index):
        obj_list = list()

        for record in child_wrapper.parentObjWrapperList[parent_index].grandObjWrapperList:
            obj_list.append(record.objName)
        if len(obj_list) > 0:
            try:
                return obj_list.index(current_obj)
            except ValueError:
                return -1
        else:
            return -1

    # Method to return index of the field obj
    def check_field_obj_present(self, current_obj, child_wrapper):
        field_list = list()

        for record in child_wrapper.fieldWrapperList:
            field_list.append(record.fieldName)
        if len(field_list) > 0:
            try:
                return field_list.index(current_obj)
            except ValueError:
                return -1
        else:
            return -1

    # Method to generate childWrapper
    def generate_child_obj(self, child_obj, child_wrapper):
        child1_field_wrap = field_wrap_obj(child_obj[2], False)
        field_index = self.check_field_obj_present(child_obj[2], child_wrapper)
        if len(child_obj) == 3:
            if field_index == -1:
                child_wrapper.fieldWrapperList.append(child1_field_wrap)
        if len(child_obj) == 4:
            if field_index == -1:
                child_wrapper.fieldWrapperList.append(child1_field_wrap)
            index_value = self.check_obj_present(child_obj[2], child_wrapper)
            if index_value == -1:
                child_wrapper.parentObjWrapperList.append(parent_wrap_obj(
                    child_obj[2], False, [field_wrap_obj(child_obj[3], False)], [], [], []))
            else:

                if field_wrap_obj(child_obj[3], False) not in child_wrapper.parentObjWrapperList[check_obj_present(child_obj[2], child_wrapper)].fieldWrapperList:
                    child_wrapper.parentObjWrapperList[check_obj_present(
                        child_obj[2], child_wrapper)].fieldWrapperList.append(field_wrap_obj(child_obj[3], False))
        if len(child_obj) == 5:
            if field_index == -1:
                child_wrapper.fieldWrapperList.append(child1_field_wrap)
            index_value = self.check_obj_present(child_obj[2], child_wrapper)
            if index_value == -1:
                child_wrapper.parentObjWrapperList.append(parent_wrap_obj(child_obj[2], False,
                                                                          [field_wrap_obj(
                                                                              child_obj[3], False)],
                                                                          [grand_wrap_obj(child_obj[3], False, [
                                                                              field_wrap_obj(child_obj[4], False)])]
                                                                          ))
            else:
                if field_wrap_obj(child_obj[3], False) not in child_wrapper.parentObjWrapperList[index_value].fieldWrapperList:
                    child_wrapper.parentObjWrapperList[index_value].fieldWrapperList.append(
                        field_wrap_obj(child_obj[3], False))
                grand_index_value = check_grand_obj_present(
                    child_obj[3], child_wrapper, index_value)
                if grand_index_value == -1:
                    child_wrapper.parentObjWrapperList[index_value].grandObjWrapperList.append(
                        grand_wrap_obj(child_obj[3], False, [field_wrap_obj(child_obj[4], False)]))
                else:
                    child_wrapper.parentObjWrapperList[index_value].grandObjWrapperList[grand_index_value].fieldWrapperList.append(
                        field_wrap_obj(child_obj[4], False))
        return child_wrapper

    def get_all_table_patterns(self, whole_text, table_pattern_list):
        table_patterns = re.findall(
            "\\$tbl\\{START:.*?\\}(.*?)\\$tbl\\{END:.*?\\}", whole_text)
        if len(table_patterns) > 0:
            table_pattern_list.append(table_patterns[0])
            remaining_text = whole_text.index('END')
            if remaining_text != -1:
                return self.get_all_table_patterns(whole_text[remaining_text+3:], table_pattern_list)
            else:
                return table_pattern_list
        else:
            return table_pattern_list

    # Method to manipulate functions in the document

    def generate_functions(self, function_list, data_dict):
        if_condition_list = re.findall("IF\\((.*?)\\)", function_list[0])
        if len(if_condition_list) > 0:
            conditon_value, true_value, false_value = if_condition_list[0].split(
                ',')[0], if_condition_list[0].split(',')[1], if_condition_list[0].split(',')[2]
            field_name_list = re.findall("\\$\\{(.*?)\\}", conditon_value)
            if '==' in str(conditon_value):
                conv_value_to_str = re.split('== ', str(conditon_value))
            if '!=' in str(conditon_value):
                conv_value_to_str = re.split('!= ', str(conditon_value))
            if '>=' in str(conditon_value):
                conv_value_to_str = re.split('>= ', str(conditon_value))
            if '<=' in str(conditon_value):
                conv_value_to_str = re.split('<= ', str(conditon_value))
            if '>' in str(conditon_value):
                conv_value_to_str = re.split('> ', str(conditon_value))
            if '<' in str(conditon_value):
                conv_value_to_str = re.split('< ', str(conditon_value))
            added_changes = conditon_value.replace(
                conv_value_to_str[-1], "'"+conv_value_to_str[-1]+"'")
            if len(field_name_list) > 0:
                splited_list = field_name_list[0].split('.')
                if len(splited_list) == 2:
                    field_value = data_dict[splited_list[1]]
                elif len(splited_list) == 3:
                    obj_name_match = re.split('Id', splited_list[1])
                    field_value = data_dict[obj_name_match[0]][splited_list[2]]
                elif len(splited_list) == 4:
                    parent_name_match = re.split('Id', splited_list[1])
                    grand_name_match = re.split('Id', splited_list[2])
                    field_value = data_dict[parent_name_match[0]
                                            ][grand_name_match[0]][splited_list[3]]
            field_value = str(field_value)
            val = added_changes.replace(
                '${'+field_name_list[0]+'}', "'"+field_value.strip()+"'")
            if bool(re.match('^(?=.*[a-zA-Z])', val)) == False:
                val = val.replace("'", "")
            cons = eval("true_value if "+val+" else false_value")
            return cons
        else:
            return "Error"

    # Method to get field index
    # Parameters (fieldName, metaData, objName)
    def get_field_index(field_name, data, list_name, obj_name):
        obj_list = list()
        for record in data[list_name]:
            obj_list.append(record[obj_name])
        if len(obj_list) > 0:
            try:
                return obj_list.index(field_name)
            except ValueError:
                return -1
        else:
            return -1

    # To bind values from salesforce to the matched string
    # Parameters(fieldName, metaData, filePath)

    def attach_field_values(self, obj_to_bind, data_dict):
        function_list = re.findall("\\{\\{FUNC:(.*?)\\}}", obj_to_bind)
        field_name = ''
        if len(function_list) > 0:
            field_name = self.generate_functions(function_list, data_dict)
        else:
            format_type = re.findall("( #[A-Z]*)", obj_to_bind)
            date_syntax = re.findall("( #DATE.*)", obj_to_bind)
            corrected_field = ''
            if len(date_syntax) > 0:
                corrected_field = obj_to_bind.replace(
                    date_syntax[0].strip(), '').strip()
            elif len(format_type) > 0:
                corrected_field = obj_to_bind.replace(
                    format_type[0].strip(), '').strip()
            else:
                corrected_field = obj_to_bind
            splited_list = corrected_field.split('.')
            if len(splited_list) == 2:
                formatted_type_data = data_dict[splited_list[1].strip()
                                                ] if splited_list[1] in data_dict.keys() else ''
                formatted_type = str(formatted_type_data)
                value = ''

                if len(format_type) > 0 and format_type[0] == ' #NUMBER':
                    value = ','.join(formatted_type[i:i+3]
                                     for i in range(0, len(formatted_type), 3))
                    formatted_type = value
                elif len(format_type) > 0 and format_type[0] == ' #CURRENCY':
                    curr_value = locale.currency(
                        formatted_type_data, grouping=True)
                    formatted_type = curr_value
                elif len(format_type) > 0 and format_type[0] == ' #DATE':
                    separate_date = formatted_type.split('-')
                    datefield = separate_date[2][:2]
                    value = date(int(separate_date[0]), int(
                        separate_date[1]), int(datefield)).ctime()
                    value = value.split(' ')
                    if len(date_syntax) > 0:
                        if re.findall('(DD/MM/YYYY)', date_syntax[0]):
                            value = datefield + '/' + \
                                separate_date[1]+'/'+separate_date[0]
                            formatted_type = value
                        elif re.findall('(DD-MM-YYYY)', date_syntax[0]):
                            value = datefield + '-' + \
                                separate_date[1]+'-'+separate_date[0]
                            formatted_type = value
                        elif re.findall('(MM-DD-YYYY)', date_syntax[0]):
                            value = separate_date[1] + '-' + \
                                datefield+'-'+separate_date[0]
                            formatted_type = value
                        elif re.findall('(MM/DD/YYYY)', date_syntax[0]):
                            value = separate_date[1] + '/' + \
                                datefield+'/'+separate_date[0]
                            formatted_type = value
                    else:
                        value = value[1]+' '+value[2]+','+''+value[-1]
                        formatted_type = value
                field_name = formatted_type
            elif len(splited_list) == 3:
                obj_name_match = re.split('Id', splited_list[1])
                try:
                    formatted_type_data = data_dict[obj_name_match[0]][splited_list[2]
                                                                       ] if splited_list[2] in data_dict[obj_name_match[0]].keys() else ''

                except KeyError:
                    formatted_type_data = ''
                formatted_type = str(formatted_type_data)
                if len(format_type) > 0 and format_type[0] == ' #NUMBER':
                    value = ','.join(formatted_type[i:i+3]
                                     for i in range(0, len(formatted_type), 3))
                elif len(format_type) > 0 and format_type[0] == ' #CURRENCY':
                    curr_value = locale.currency(
                        formatted_type_data, grouping=True)
                    value = curr_value
                elif len(format_type) > 0 and format_type[0] == ' #DATE':
                    if formatted_type != '':
                        separate_date = formatted_type.split('-')
                        datefield = separate_date[2][:2]
                        value = date(int(separate_date[0]), int(
                            separate_date[1]), int(datefield)).ctime()
                        value = value.split(' ')
                        value = value[1]+' '+value[2]+','+''+value[-1]
                    else:
                        value = obj_to_bind
                field_name = formatted_type
            elif len(splited_list) == 4:
                obj_name_match = re.split('Id', splited_list[1])
                obj_field_name = re.split('Id', splited_list[2])
                try:
                    formatted_type_data = data_dict[obj_name_match[0]
                                                    ][obj_field_name[0]][splited_list[3]]
                except KeyError:
                    formatted_type_data = ''
                formatted_type = str(formatted_type_data)
                #  formatted_type = data_dict[parent_name_match[0]][grand_name_match[0]][splited_list[3]] if [splited_list[3]] in parent_list.keys() else ''
                if len(format_type) > 0 and format_type[0] == ' #NUMBER':
                    value = ','.join(formatted_type[i:i+3]
                                     for i in range(0, len(formatted_type), 3))
                elif len(format_type) > 0 and format_type[0] == ' #CURRENCY':
                    curr_value = locale.currency(
                        formatted_type_data, grouping=True)
                    value = curr_value
                elif len(format_type) > 0 and format_type[0] == ' #DATE':
                    separate_date = formatted_type.split('-')
                    datefield = separate_date[2][:2]
                    value = date(int(separate_date[0]), int(
                        separate_date[1]), int(datefield)).ctime()
                    value = value.split(' ')
                    value = value[1]+' '+value[2]+','+''+value[-1]
                field_name = formatted_type
        return field_name

    def form_child_obj_fields(self, child_obj_metadata, pick_list_fields, child_table_list):
        obj_wrapper_list = []
        obj_wrapper_child = type('test', (object,), {})()
        for child_obj in child_table_list:
            obj_wrapper_child = obj_wrap(child_obj, False, [], [], [])
            field_wrapper = []
            parent_wrapper = []
            parent_field_wrapper = []
            grand_parent_field_wrapper = []
            grand_wrapper = []
            for field in child_obj_metadata:
                if re.search(child_obj, field.split('.')[0]):
                    date_type = re.findall("( #DATE.*)", field)
                    format_type = re.findall("( #[A-Z]*)", field)
                    if len(format_type) > 0 and format_type[0] == ' #PICKLIST':
                        corrected_field = field.replace(
                            format_type[0].strip(), '').strip()
                        parent_obj = corrected_field.split('.')
                        pick_list_fields.append(parent_obj[-1])
                    elif len(date_type) > 0:
                        corrected_field = field.replace(
                            date_type[0].strip(), '').strip()
                    elif len(format_type) > 0:
                        corrected_field = field.replace(
                            format_type[0].strip(), '').strip()
                    else:
                        corrected_field = field

                    parent_obj = corrected_field.split('.')
                    if len(parent_obj) == 2:
                        field_wrap = field_wrap_obj(parent_obj[-1], False)
                        if field_wrap not in field_wrapper:
                            field_wrapper.append(field_wrap)

                    else:
                        field_wrap = field_wrap_obj(parent_obj[1], False)
                        if field_wrap.__dict__ not in field_wrapper:
                            field_wrapper.append(field_wrap.__dict__)
                        filtered_obj = parent_obj[1:len(parent_obj)]

                        if len(filtered_obj) == 3:
                            parent_field_wrap = field_wrap_obj(
                                filtered_obj[1], False)
                            if parent_field_wrap not in parent_field_wrapper:
                                parent_field_wrapper.append(parent_field_wrap)
                            grand_parent_field_wrap = field_wrap_obj(
                                filtered_obj[2], False)
                            if grand_parent_field_wrap not in grand_parent_field_wrapper:
                                grand_parent_field_wrapper.append(
                                    grand_parent_field_wrap)
                            grand_wrap = grand_wrap_obj(
                                filtered_obj[1], False, grand_parent_field_wrapper)
                            if len(grand_wrapper) > 0:
                                check_grobj_list = list()
                                for obj in grand_wrapper:
                                    check_grobj_list.append(obj.objName)

                                if grand_wrap.objName not in check_grobj_list:
                                    grand_wrapper.append(grand_wrap)

                                elif {
                                        'fieldName': filtered_obj[2]
                                } not in grand_wrapper[check_grobj_list.index(
                                        grand_wrap.objName)].fieldWrapperList:
                                    grand_wrapper[check_grobj_list.index(
                                        grand_wrap.objName)].fieldWrapperList.append(
                                            field_wrap_obj(filtered_obj[2], False))
                            else:
                                grand_wrapper.append(grand_wrap)
                            parent_wrap = parent_wrap_obj(
                                filtered_obj[0], False, parent_field_wrapper, [], [], grand_wrapper)

                            if len(parent_wrapper) > 0:
                                check_obj_list = list()
                                for obj in parent_wrapper:
                                    check_obj_list.append(obj.objName)
                                if parent_wrap.objName not in check_obj_list:
                                    parent_wrapper.append(parent_wrap)

                                else:

                                    if field_wrap_obj(filtered_obj[1], False) not in parent_wrapper[check_obj_list.index(
                                            parent_wrap.objName)].fieldWrapperList:
                                        parent_wrapper[check_obj_list.index(
                                            parent_wrap.objName
                                        )].fieldWrapperList.append(field_wrap_obj(filtered_obj[1], False))
                                        parent_wrapper[check_obj_list.index(
                                            parent_wrap.objName
                                        )].grandObjWrapperList.append(grand_wrap_obj(filtered_obj[1], False, [field_wrap_obj(filtered_obj[2], False)]))
                                    else:

                                        check_grandobj_list = list()
                                        if 'grandObjWrapperList' in parent_wrapper[
                                                check_obj_list.index(
                                                    parent_wrap.objName)]:

                                            for obj in parent_wrapper[check_obj_list.index(
                                                    parent_wrap.objName
                                            )].grandObjWrapperList:
                                                check_grandobj_list.append(
                                                    obj.objName)
                                        else:
                                            check_grandobj_list = []
                                        if grand_wrap.objName not in check_grandobj_list:
                                            grand_wrapper.append(grand_wrap)
                                        else:
                                            if field_wrap_obj(filtered_obj[2], False) not in parent_wrapper[check_obj_list.index(
                                                    parent_wrap.objName
                                            )].grandObjWrapperList[
                                                    check_grandobj_list.index(
                                                        grand_wrap.objName
                                                    )].fieldWrapperList:

                                                parent_wrapper[check_obj_list.index(
                                                    parent_wrap.objName
                                                )].grandObjWrapperList[
                                                    check_grandobj_list.index(
                                                        grand_wrap.objName
                                                    )].fieldWrapperList.append(field_wrap_obj(filtered_obj[2], False))

                            else:
                                parent_wrapper.append(parent_wrap)

                            grand_wrap = {}
                            grand_parent_field_wrap = {}
                            parent_field_wrap = {}
                            parent_wrap = {}
                            parent_field_wrapper = []
                            grand_parent_field_wrapper = []
                            grand_wrapper = []

                        elif len(filtered_obj) == 2:
                            parent_field_wrap = field_wrap_obj(
                                filtered_obj[-1], False)
                            check_parent_obj_list = list()
                            for obj in parent_wrapper:
                                check_parent_obj_list.append(obj.objName)
                            if ({
                                    'objName': filtered_obj[0]
                            } in check_parent_obj_list
                            ) and parent_field_wrap not in parent_wrapper[
                                    check_parent_obj_list.index(
                                        filtered_obj[0])]['fieldWrapperList']:
                                parent_field_wrapper.append(parent_field_wrap)

                            parent_wrap = parent_wrap_obj(
                                filtered_obj[0], False, parent_field_wrapper, [], [], [])
                            if len(parent_wrapper) > 0:
                                check_obj_list = list()
                                for obj in parent_wrapper:
                                    check_obj_list.append(obj.objName)
                                if parent_wrap.objName not in check_obj_list:
                                    parent_wrap.fieldWrapperList = [
                                        field_wrap_obj(filtered_obj[-1], False)]
                                    parent_wrapper.append(parent_wrap)
                                    parent_field_wrapper = []

                                elif {
                                        'fieldName': filtered_obj[-1],
                                        'isExists': False
                                } not in parent_wrapper[check_obj_list.index(
                                        parent_wrap.objName)].fieldWrapperList:
                                    parent_wrapper[check_obj_list.index(
                                        parent_wrap.objName
                                    )].fieldWrapperList.append(field_wrap_obj(filtered_obj[-1], False))
                                    parent_field_wrapper = []
                            else:
                                parent_wrap.fieldWrapperList.append(
                                    field_wrap_obj(filtered_obj[-1], False))
                                parent_wrapper.append(parent_wrap)
                                parent_field_wrapper = []
                        parent_field_wrapper = []

                obj_wrapper_child.fieldWrapperList = field_wrapper
                obj_wrapper_child.parentObjWrapperList = parent_wrapper
            obj_wrapper_list.append(obj_wrapper_child)
        return obj_wrapper_list
