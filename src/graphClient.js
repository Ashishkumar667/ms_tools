const axios = require("axios");

/**
 * Simple Microsoft Graph client that wraps axios.
 * You pass in a valid access token (acquired via Azure AD, client credentials, etc.).
 */
class GraphClient {
  /**
   * @param {string} accessToken - A bearer token for Microsoft Graph.
   * @param {object} [options]
   * @param {string} [options.baseUrl] - Graph base URL, defaults to v1.0.
   */
  constructor(accessToken, options = {}) {
    if (!accessToken) {
      throw new Error("accessToken is required to initialize GraphClient");
    }

    this.accessToken = accessToken;
    this.baseUrl = options.baseUrl || "https://graph.microsoft.com/v1.0";

    this.http = axios.create({
      baseURL: this.baseUrl,
      headers: {
        Authorization: `Bearer ${this.accessToken}`,
        "Content-Type": "application/json"
      },
      timeout: options.timeout || 15000
    });
  }

  /**
   * Generic request helper.
   * @param {string} method - HTTP method (GET, POST, PATCH, DELETE).
   * @param {string} path - Graph path starting with `/`, e.g. `/teams/{id}`.
   * @param {object} [data] - Request body.
   * @param {object} [config] - Extra axios config.
   */
  async request(method, path, data, config = {}) {
    const response = await this.http.request({
      method,
      url: path,
      data,
      ...config
    });
    return response.data;
  }

  get(path, config) {
    return this.request("GET", path, undefined, config);
  }

  post(path, data, config) {
    return this.request("POST", path, data, config);
  }

  patch(path, data, config) {
    return this.request("PATCH", path, data, config);
  }

  delete(path, config) {
    return this.request("DELETE", path, undefined, config);
  }
}

module.exports = {
  GraphClient
};


