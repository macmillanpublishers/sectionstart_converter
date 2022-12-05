FROM python:3.9.15-alpine3.17

WORKDIR /tmp

RUN pip install lxml requests six

# RUN apk update 

RUN apk add git diffutils

RUN git clone https://github.com/macmillanpublishers/RSuite_Word-template.git /opt/RSuite_Word-template

RUN mkdir /opt/rsv

COPY ./xml_docx_stylechecks /opt/rsv/xml_docx_stylechecks
COPY ./test /opt/rsv/test

# CMD [ "python", "./your-daemon-or-script.py" ]
# ENTRYPOINT ["top", "-b"]