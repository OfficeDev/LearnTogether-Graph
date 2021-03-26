export function CacheMiddleware(expirationConfig) {
  this.nextMiddleware = undefined;
  this.expirationConfig = expirationConfig;

  // converts headers to a hashtable for easier serializing
  const getHeaders = (headers) => {
    const h = {};
    for (var header of headers.entries()) {
      h[header[0]] = header[1];
    }
    return h;
  };

  // converts blob to a data URL for serializing
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

  // converts data URL to a blob so that it can be set on a response
  const dataUrlToBlob = (dataUrl) => {
    const blobData = dataUrl.split(',');
    const contentType = blobData[0].replace('data:', '').replace(';base64', '');
    return b64toBlob(blobData[1], contentType);
  };
  
  // from https://stackoverflow.com/a/16245768
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

  // gets expiration date for the specified URL
  // if expiration config has not been set, or the particular path has not been
  // defined, returns undefined which means that the response will be cached
  // for the duration of the session
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

      // create a cache key. We should consider including headers in the key
      // as well
      const requestKey = btoa(context.request);
      let response = window.sessionStorage.getItem(requestKey);
      let now = new Date();

      if (response) {
        const resp = JSON.parse(response);
        // check if the cached response has an expiration date configured and
        // if so, if it's still in the future (cache valid)
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

      // no cache hit or cache expired, retrieve data from Graph
      console.debug('-- from Graph');
      await this.nextMiddleware.execute(context);

      // don't cache non-GET or failed requests
      if (context.options.method !== 'GET' ||
        context.response.status !== 200) {
        return;
      }

      // clone response so that we can get the body stream
      const resp = context.response.clone();

      const expiresOn = getExpiresOn(resp.url);
      // reset date to catch expiration set to 0
      now = new Date();
      // don't cache response if it's already expired
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