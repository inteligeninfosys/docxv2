FROM node:18-alpine3.15

RUN mkdir -p /app/nfs/demandletters
RUN mkdir -p /app/docxletters
WORKDIR /app/docxletters

RUN chown node:node -R /app
USER node
# Install app dependencies
COPY --chown=node package*.json ./
RUN npm install --production
COPY --chown=node . .

EXPOSE 8040

CMD ["node", "index.js"]

#  docker build -t migutak/docxv2:5.0 .

