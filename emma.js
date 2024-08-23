const http = require('https')

httprequest().then((data) => {
        const response = {
            statusCode: 200,
            body: JSON.stringify(data),
        };
    return response;
});
function httprequest(host, path, port = 8080) {
     return new Promise((resolve, reject) => {
        const options = {
            host: host,
            path: path,
            port: port,
            method: 'GET'
        };
        const req = http.request(options, (res) => {
          if (res.statusCode < 200 || res.statusCode >= 300) {
                return reject(new Error('statusCode=' + res.statusCode));
            }
            var body = [];
            res.on('data', function(chunk) {
                body.push(chunk);
            });
            res.on('end', function() {
                try {
                    body = JSON.parse(Buffer.concat(body).toString());
                } catch(e) {
                    reject(e);
                }
                resolve(body);
            });
        });
        req.on('error', (e) => {
          reject(e.message);
        });
        // send the request
       req.end();
    });
}

console.log(httprequest('google.com', ''));
