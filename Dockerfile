FROM node:16

COPY . /root/graph-azure-ks

WORKDIR /root/graph-azure-ks

ARG NPM_ACCESS_TOKEN

RUN echo "//registry.npmjs.org/:_authToken=${NPM_ACCESS_TOKEN}" > .npmrc && \
    echo "@keystone-labs:registry=https://registry.npmjs.org" >> .npmrc && \
    yarn set version 1.22.19 && \
    yarn install && \
    yarn build && \
    rm -f .npmrc

ENTRYPOINT [ "yarn", "start" ]