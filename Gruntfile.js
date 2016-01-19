module.exports = function(grunt) {
    grunt.initConfig({
        pkg: grunt.file.readJSON('package.json'),
        concat: {
            options: {
                separator: '\n'
            },
            dist: {
                src: ['js/*'],
                dest: 'SharePoint.CustomUtilities.js'
            }
        },       
        watch: {         
            scripts: {
                files: ['js/*.js'],
                options: {
                    interrupt: true
                },
                tasks: ['concat']
            }
        }
    });
    grunt.loadNpmTasks('grunt-contrib-concat');    
    grunt.loadNpmTasks('grunt-contrib-watch');
    grunt.registerTask('default', ['watch']);
};