FROM python:3.9.15-alpine3.17

WORKDIR /tmp

RUN pip install lxml requests six

# RUN apk update

RUN apk add git diffutils

RUN mkdir /opt/rsv

COPY ./xml_docx_stylechecks /opt/rsv/xml_docx_stylechecks
COPY ./test /opt/rsv/test
COPY ./RSuite_Word-template /opt/rsv/RSuite_Word-template

# CMD [ "python", "./your-daemon-or-script.py" ]
# ENTRYPOINT ["top", "-b"]
