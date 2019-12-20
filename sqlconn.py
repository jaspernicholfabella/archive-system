from sqlalchemy import create_engine
from sqlalchemy import Table,Column,VARCHAR,INTEGER,Float,String,MetaData,ForeignKey,Date,Text,DECIMAL,Boolean
from sqlalchemy.sql import exists
import os
class Database():
    engine = create_engine('sqlite:///{}/db/archive_system.db'.format(os.getcwd()))
    meta = MetaData()

    archive_admin = Table('archive_admin',meta,
                          Column('userid',INTEGER,primary_key=True),
                          Column('username',VARCHAR(50)),
                          Column('password',VARCHAR(50)),
                          Column('previlage',VARCHAR(50)))

    archive_doctype = Table('archive_doctype',meta,
                            Column('doctype_id',INTEGER,primary_key=True),
                            Column('document_type',VARCHAR(50)))

    archive_sharedrive = Table('archive_sharedrive',meta,
                               Column('sdid',INTEGER,primary_key=True),
                               Column('sharedrive',VARCHAR(50)))

    archive_document = Table('archive_document',meta,
                             Column('docid',INTEGER,primary_key=True),
                             Column('docname',VARCHAR(50)),
                             Column('isconfidential',Boolean,unique=False,default=False),
                             Column('doctype',VARCHAR(50)),
                             Column('description',Text),
                             Column('alias',Text),
                             Column('iseditable',Boolean,unique=False,default=False),
                             Column('filetype',VARCHAR(5)),
                             Column('date_uploaded',Date))

    archive_mail = Table('archive_mail',meta,
                         Column('mailid',INTEGER,primary_key = True),
                         Column('sender',VARCHAR(50)),
                         Column('reciever',VARCHAR(50)),
                         Column('date_sent',Date),
                         Column('from_who',VARCHAR(30)),
                         Column('subject',Text),
                         Column('action',Text),
                         Column('have_attached',Boolean,unique=False,default=False),
                         Column('attached_alias',VARCHAR(100)),
                         Column('isseen',Boolean,unique=False,default=False),
                         Column('iseditable',Boolean,unique=False,default=False),
                         Column('filetype', VARCHAR(5)),
                         Column('status',VARCHAR(50)),
                         Column('status_message',VARCHAR(100)),
                         Column('reply_have_attached',Boolean,unique=False,default=False),
                         Column('reply_attached_alias',VARCHAR(100)),
                         Column('reply_is_editable', Boolean, unique=False, default=False),
                         Column('reply_filetype', VARCHAR(5)),
                         )


    meta.create_all(engine)

    conn = engine.connect()
    s = archive_sharedrive.select()
    s_value = conn.execute(s)
    x = 0
    for val in s_value:
        x += 1
    loc = os.getcwd() + r'\archive_directory'
    if x == 0:
        ins = archive_sharedrive.insert().values(sharedrive = os.getcwd())
        result = conn.execute(ins)