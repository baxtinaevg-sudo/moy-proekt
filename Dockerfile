FROM nginx:alpine
# Force rebuild v2
RUN echo "rebuild-v3" > /tmp/version
COPY index.html /usr/share/nginx/html/index.html
EXPOSE 80
