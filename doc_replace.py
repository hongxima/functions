# -*- coding: utf-8 -*-
import logging
import json
import oss2
import uuid
from docx import Document
from python_docx_replace import docx_replace


def handler(event, context):
  evt = json.loads(event)
  print (u"event = {}".format(evt))
  values = json.loads(evt.get('body'))

  new_filename = u"{}.docx".format(uuid.uuid4())
  output_path = u"/tmp/{}".format(new_filename)
  template_filename = u"/tmp/doc-template.docx"
  
  try:  
    auth = oss2.StsAuth(context.credentials.access_key_id, context.credentials.access_key_secret, context.credentials.security_token)
    endpoint = 'https://oss-cn-shanghai-internal.aliyuncs.com'
    bucket = oss2.Bucket(auth, endpoint, 'nbagent-testdata')
    bucket.get_object_to_file('doc-template-v1.docx', template_filename)
    

    template = Document(template_filename)
    docx_replace(template, **values)
    template.save(output_path)
    print(f"Creation successful. Word document saved at: {output_path}")

    bucket.put_object_from_file(new_filename, output_path)
    return u"https://nbagent-testdata.oss-cn-shanghai.aliyuncs.com/{}".format(new_filename)
  except Exception as e:
    print(f"Error during docx creation: {e}")
