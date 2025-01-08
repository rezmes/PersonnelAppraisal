"use strict";

// const build = require('@microsoft/sp-build-web');

// build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

// build.initialize(require('gulp'));
"use strict";

const build = require("@microsoft/sp-build-web");
const webpack = require("webpack");

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    generatedConfiguration.resolve.fallback = {
      crypto: require.resolve("crypto-browserify"),
      buffer: require.resolve("buffer"),
      stream: require.resolve("stream-browserify"),
      assert: require.resolve("assert"),
      util: require.resolve("util"),
      process: require.resolve("process/browser"),
    };

    generatedConfiguration.plugins.push(
      new webpack.ProvidePlugin({
        Buffer: ["buffer", "Buffer"],
        process: "process/browser",
      })
    );

    return generatedConfiguration;
  },
});

build.initialize(require("gulp"));
