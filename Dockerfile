FROM alpine:latest AS alpine

FROM docker.n8n.io/n8nio/n8n:2.32.6
COPY --from=alpine /sbin/apk /sbin/apk
COPY --from=alpine /usr/lib/libapk.so* /usr/lib/
COPY --from=alpine /lib/apk /lib/apk
COPY --from=alpine /etc/apk /etc/apk

USER root
RUN apk add --no-cache python3 py3-pip && \
    pip3 install --break-system-packages openpyxl
COPY master_consolidator.py /opt/scripts/master_consolidator.py
USER node
