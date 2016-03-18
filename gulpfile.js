/// Gulp configuration
'use strict';

var gulp = require('gulp'),
    connect = require('gulp-connect'),
    config = require('./gulpfile.config.json');

gulp.task('default', function() {
    return connect.server({
        root: config.server.root,
        host: config.server.host,
        port: config.server.port,
        livereload: true,
        https: true,
    });
});
