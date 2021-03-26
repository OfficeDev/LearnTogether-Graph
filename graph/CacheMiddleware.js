export function CacheMiddleware(expirationConfig) {
  this.nextMiddleware = undefined;
  this.expirationConfig = expirationConfig;

  const getHeaders = (headers) => {
    const h = {};
    for (var header of headers.entries()) {
      h[header[0]] = header[1];
    }
    return h;
  };

  const blobToDataUrl = async (blob) => {
    return new Promise((resolve, reject) => {
      var reader = new FileReader();
      reader.onload = function () {
        var dataUrl = reader.result;
        resolve(dataUrl);
      };
      reader.readAsDataURL(blob);
    });
  };

  // from https://stackoverflow.com/a/16245768
  const dataUrlToBlob = (dataUrl) => {
    const blobData = dataUrl.split(',');
    const contentType = blobData[0].replace('data:', '').replace(';base64', '');
    return b64toBlob(blobData[1], contentType);
  };

  const b64toBlob = (b64Data, contentType = '', sliceSize = 512) => {
    const byteCharacters = atob(b64Data);
    const byteArrays = [];

    for (let offset = 0; offset < byteCharacters.length; offset += sliceSize) {
      const slice = byteCharacters.slice(offset, offset + sliceSize);

      const byteNumbers = new Array(slice.length);
      for (let i = 0; i < slice.length; i++) {
        byteNumbers[i] = slice.charCodeAt(i);
      }

      const byteArray = new Uint8Array(byteNumbers);
      byteArrays.push(byteArray);
    }

    const blob = new Blob(byteArrays, { type: contentType });
    return blob;
  }

  const getExpiresOn = (url) => {
    if (!this.expirationConfig) {
      return undefined;
    }

    for (var i = 0; i < this.expirationConfig.length; i++) {
      const exp = this.expirationConfig[i];
      if (url.indexOf(exp.path) > -1) {
        const expiresOn = new Date();
        expiresOn.setMinutes(expiresOn.getMinutes() + exp.expirationInMinutes);
        return expiresOn;
      }
    }

    return undefined;
  }

  return {
    execute: async (context) => {
      console.debug(`Request: ${context.request}`);

      const requestKey = btoa(context.request);
      let response = window.sessionStorage.getItem(requestKey);
      let now = new Date();
      if (response) {
        const resp = JSON.parse(response);
        const expiresOn = resp.expiresOn ? new Date(resp.expiresOn) : undefined;
        if (!expiresOn || expiresOn > now) {
          console.debug('-- from cache');

          let body;
          if (resp.headers['content-type'].indexOf('application/json') > -1) {
            body = JSON.stringify(resp.body);
          }
          else {
            body = dataUrlToBlob(resp.body);
          }
          context.response = new Response(body, resp);
          return;
        }
        else {
          console.log('-- cache expired');
        }
      }

      console.debug('-- from Graph');
      await this.nextMiddleware.execute(context);

      if (context.options.method !== 'GET' ||
        context.response.status !== 200) {
        // don't cache non-GET or failed requests
        return;
      }

      const resp = context.response.clone();

      const expiresOn = getExpiresOn(resp.url);
      // reset date to catch expiration set to 0
      now = new Date();
      // don't cache if the item already expired
      if (expiresOn <= now) {
        return;
      }

      const headers = getHeaders(resp.headers);
      let body = '';
      if (headers['content-type'].indexOf('application/json') >= 0) {
        body = await resp.json();
      }
      else {
        body = await blobToDataUrl(await resp.blob());
      }

      response = {
        url: resp.url,
        status: resp.status,
        statusText: resp.statusText,
        headers,
        body,
        expiresOn
      };
      window.sessionStorage.setItem(requestKey, JSON.stringify(response));
    },
    setNext: (next) => {
      this.nextMiddleware = next;
    }
  }
};