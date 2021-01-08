from bson.objectid import ObjectId
import re


class generate_mongo_query:
    def form_mongo_query(self, json_dic, record_id, pick_list_fields, where_condition):
        mongodbquery = []
        final_projection_list = []
        parent_object_list = []
        objMatch = {
            '$match': {
                where_condition: ObjectId(record_id)
            }
        }
        mongodbquery.append(objMatch)
        final_projection_list.append('_id')
        for fields_in_mainobject in json_dic.get('fieldWrapperList'):
            field_api_name = ''
            main_obj_api = fields_in_mainobject['fieldName'].split('(')
            field_api_name = main_obj_api[0]
            final_projection_list.append(fields_in_mainobject['fieldName'])
            pass
        if len(json_dic['parentObjWrapperList']) > 0:
            for parent_obj_list in json_dic.get('parentObjWrapperList'):
                #  get field name and api name from the parent object
                parent_obj_api = parent_obj_list.get('objName').split('(')
                obj_name = parent_obj_api[1].split(')')[0]
                field_name = parent_obj_api[0]
                field_query_list = []
                org_name = json_dic['objName'].split('#')[1]
                # get all fields in the variable
                for fields in parent_obj_list.get('fieldWrapperList'):
                    field_query_list.append(fields.get('fieldName'))
                field_query = self.Convert(field_query_list)
                parent_object_list.append(parent_obj_list.get('objName'))
                parent_obj = {
                    "$lookup": {
                        "from": org_name+'.'+obj_name,
                        "let": {
                            "id": '$'+field_name
                        },
                        "pipeline": [
                            {
                                "$match": {
                                    "$expr": {
                                        "$eq": [
                                            "$_id",
                                            "$$id"
                                        ]
                                    }
                                }
                            },
                            {
                                "$project": field_query
                            }
                        ],
                        "as": parent_obj_list.get('objName')
                    }
                }
                mongodbquery.append(parent_obj)
        if len(json_dic.get('childObjWrapperList')) > 0:
            print('childObjWrapperList-->', json_dic.get('childObjWrapperList'))
            for child_obj_value in json_dic.get('childObjWrapperList'):
                child_obj_api = child_obj_value.get('objName').split('(')
                field_parent_name = child_obj_api[1].split(')')[0]
                obj_child_name = child_obj_api[0]
                field_query_list = []
                # get all fields in the variable
                for fields in child_obj_value.get('fieldWrapperList'):
                    field_query_list.append(fields.get('fieldName'))
                field_query = self.Convert(field_query_list)
                parent_obj = {
                    "$lookup": {
                        "from": obj_child_name.split('#')[1]+'.'+obj_child_name.split('#')[2],
                        "let": {
                            "id": "$_id"
                        },
                        "pipeline": [
                            {
                                "$match": {
                                    "$expr": {
                                        "$eq": [
                                            '$'+field_parent_name,
                                            "$$id"
                                        ]
                                    }
                                }
                            },
                            {
                                "$project": field_query
                            }
                        ],
                        "as": child_obj_value.get('objName')
                    }
                }
                mongodbquery.append(parent_obj)
                final_projection_list.append(child_obj_value.get('objName'))

        all_object_projection = self.Convert(final_projection_list)
        picklist_collection_obj = []
        for key, value in all_object_projection.items():
            if key in parent_object_list:
                all_object_projection[key] = {
                    "$arrayElemAt": [
                        "$"+key,
                        0
                    ]
                }
            if key in pick_list_fields:
                all_object_projection[key] = {
                    "$arrayElemAt": [
                        "$"+key+'.label',
                        0
                    ]
                }
                picklist_collection_obj.append({
                    "$lookup": {
                        "from": "Picklist",
                        "let": {
                            "id": "$"+key
                        },
                        "pipeline": [
                            {
                                "$match": {
                                    "$expr": {
                                        "$eq": [
                                            "$_id",
                                            "$$id"
                                        ]
                                    }
                                }
                            },
                            {
                                "$project": {
                                    "label": 1
                                }
                            }
                        ],
                        "as": key
                    }
                })
        for pick_obj in picklist_collection_obj:
            mongodbquery.append(pick_obj)
        final_project = {
            "$project": all_object_projection
        }
        mongodbquery.append(final_project)
        return mongodbquery

    def Convert(self, lst):
        res_dct = {lst[i]: 1 for i in range(0, len(lst), 1)}
        return res_dct

    def generate_picklist_query(self, data, pick_list_fields_in_child_obj, db):
        picklist_recordids = []
        for obj_field in data:
            if type(data[obj_field]) in (tuple, list):
                for arrrecord in data[obj_field]:
                    for field_value in pick_list_fields_in_child_obj:
                        if arrrecord[field_value['field']]:
                            if arrrecord[field_value['field']] not in picklist_recordids:
                                picklist_recordids.append(
                                    arrrecord[field_value['field']])

        picklist_query = [
            {
                "$match": {
                    "_id": {
                        "$in": picklist_recordids
                    }
                }
            },
            {
                "$project": {
                    "_id": 1,
                    "label": 1
                }
            }
        ]
        collection_name = db['Picklist']
        queried_data = collection_name.aggregate(picklist_query)

        record_data = list(queried_data)
        new_data = record_data[0]
        for obj_field in data:
            if type(data[obj_field]) in (tuple, list):
                for arrrecord in data[obj_field]:
                    for field_value in pick_list_fields_in_child_obj:
                        for pick in record_data:
                            if re.search(str(pick['_id']), str(arrrecord[field_value['field']])):
                                arrrecord[field_value['field']] = pick['label']
        return data
