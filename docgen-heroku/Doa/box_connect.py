from boxsdk import Client
from boxsdk import JWTAuth
from boxsdk.exception import BoxAPIException
import traceback
import json


class logging_box:
    def connect_box(self):
        auth = JWTAuth(
            client_id='fj5klf35gaydgnx2hewpfdg4uqz85ri7',
            client_secret='mM2d0ZScYgg6dE22jShBgK6y5xL6xFID',
            enterprise_id='635964596',
            jwt_key_id='061jkz4x',
            rsa_private_key_file_sys_path=None,
            rsa_private_key_data='-----BEGIN ENCRYPTED PRIVATE KEY-----\nMIIFDjBABgkqhkiG9w0BBQ0wMzAbBgkqhkiG9w0BBQwwDgQICUFlpOqJ2DkCAggA\nMBQGCCqGSIb3DQMHBAhk2ZclZTqcNASCBMhSLnADIn+MhR1PtS0YWLej0bUrIxjE\n7uzrtYwXN9b8G7Y7PUhroEetlRIS5Uf4qEDsYFJeZwlW/wd9Xzzjdhg+bUMmVLC5\n57tcPnS9Sv0Q3gDcGiJrhjm9k7u17/GgHhCOeUeakZX8U/RJfasijw5X8ue8+esY\ng+99BB83H+9anSjTYqic2dU7PtwDdx9AYeru6g1psBjZurzT0eE1SWdTOw5QGhG4\n78+gBsuzuU4D2SefS1wMFiuY45JvxUq6+zAGtzRi8MYXYruRsi3vQRkctTqcJZil\n6g7yxs/oY0xaSNocMY7y8kuHVirdYzj3KZX/+p8EvVB+1pzVJ50dFtJntzdpcu03\nfCfLYz2Vod6TXnGs2Dr6Zzq9NSb5otumvn6FrD5QeAoVKkEcakPbGOLACBKCJe7B\nfoDB2FJXC2LqZwDsf0X7kLcJh7xaf6l2HOOp73wsZUfpifGmmswmL3dfEbZ3uLeR\nIBIdSlNyWuwrOuNc+5fl0V1zBdys3RGo4pavwq7QDFVSg3FgwtAqfOFICUlO48EL\nuF9Z3xp8pz+rlfYiF4iMpbIqt1/OD6GOFH2PiVIUaf4xB7Arbbbv93JKnTzwj8j1\nbKowjITWVc4RmPI0XbCVvukAYm4hY1ohlTaFVs5+3Dg6BQTqEtkfxknCphtPUoGD\nKV+c0YxI4jz9bYLPAWsMvwlvfYg10kUleyzKZAGt+NvBeplbi73Je7as1ZcvNcJX\nJ1jBLatH24RB7XU0N2r0KjIKMaE7pMDghTjtoJ1DczIi2t+J5dymwqAyoba0FktF\nA+T/AsAfjs95ashgekL0g0odKeCQL6YOnb41AKFZ9m6aOOONF0q8S4oLKYdqoL5U\nWq/2DnMyiKNH7eZfbAKMAJzYDjjkVKXj2MBDg/tF4ibwoVoa6jJZT6eKi82moCr5\n3LPQpJsIlfQ5nZg6faito35KlKEUCLQkSPP9p+sAZ3qyqirvkzps1c2FRyuUnVWH\nCMPUKGHcOhdq/2ziSLS8a1X95PxpWHCOS4BRnHEiGX0m2WleqkiQsJRnXrAw8Ft8\neu6knlL1o51EE4EaQeipfHaBxRDhdLMxEJGmj4mOTOP9Q8It9Sk4FNu8Sf3bXV85\nBpK30tBHFgSo04DQdxFrcHykEuFR0CjEf26us6Thp8YkAov/ogQb6EdjaYUQpm2P\nGdWxle14L/QAhkzXGA3cVng5bbeqU3gdagCe+UN1gCWGXer6OFmA3jjy5jZCrlDM\nd1O2HonoFdlh0V+RhJtvJWi2P/KcIOu/LroBnR0wcOles/TqnSHZ3LvPp2f+SfP+\nL2KL+jgZR24vofh+owEELYSD2luoi1H7X79cCWP6qxfYBj7jKzUBUkVjKOZ5QwIo\nlpCspT9mHzDB+K83p9nwwrnYy+J55InT0tiLtWvpN3og6kPy+vdX9SN1rp0VImxO\n6n1pfzYwfbhGLgc953/avlU9vSSHjAA9wFsONB344H+QMT1pgMh6D4O42XfZzriY\ne/rgz/EH7v39RQL4Y4VJbmdv9WW3NqBfLGwGgVuGFhh0wSc3ct/PYImeiijYviiD\nOvzIFZexsc5jgms0qr7Iby2fqYAwWYhxGN/uqpfDq0p9TYti7XUNHsDa0jAQrEb/\no1A=\n-----END ENCRYPTED PRIVATE KEY-----\n',
            rsa_private_key_passphrase='b8560c4b0187d6d5a068a637cfd034a9'
        )
        try:
            return Client(auth)
        except BoxAPIException as error:
            error_obj = {"isSuccess": False, "message":  str(
                error), "traceback": traceback.format_exc()}
            return json.dumps(error_obj)

    def insert_file(self, boxClient, folder_id, file_dir):
        try:
            new_file = boxClient.folder(folder_id).upload(file_dir)
            return 'File "{0}" uploaded to Box with file ID {1}'.format('testappgen1.docx', new_file.id)
        except TypeError as error:
            error_obj = {"isSuccess": False, "message":  str(
                error), "traceback": traceback.format_exc()}
            return json.dumps(error_obj)

    def get_file_content(self, boxClient, file_id):
        try:
            return boxClient.file(file_id=file_id).content()
        except TypeError as error:
            error_obj = {"isSuccess": False, "message":  str(
                error), "traceback": traceback.format_exc()}
            return json.dumps(error_obj)
