export function CacheMiddleware() {
  this.nextMiddleware = undefined;
  
  const getHeaders = (headers) => {
    const h = {};
    for (var header of headers.entries()) {
      h[header[0]] = header[1];
   }
   return h;
  };

  return {
    execute: async (context) => {
      const requestKey = btoa(context.request);
      const response = window.localStorage.getItem(requestKey);
      if (response) {
        const resp = JSON.parse(response);
        context.response = new Response(resp.body, resp);
      }
      else {
        await this.nextMiddleware.execute(context);
        const resp = context.response.clone();
        const response = {
          url: resp.url,
          status:  resp.status,
          statusText:  resp.statusText,
          headers: getHeaders(resp.headers),
          body: await resp.json()
        };
        window.localStorage.setItem(requestKey, JSON.stringify(response));
      }
    },
    setNext: (next) => {
      this.nextMiddleware = next;
    }
  }
};