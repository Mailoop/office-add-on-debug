
module.exports = function override(config, env) {
  config.entry = ["babel-polyfill","whatwg-fetch", "./src/index.js"];
  return config
};