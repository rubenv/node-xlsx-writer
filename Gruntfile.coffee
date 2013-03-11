module.exports = (grunt) ->
    @loadNpmTasks('grunt-contrib-clean')
    @loadNpmTasks('grunt-contrib-coffee')
    @loadNpmTasks('grunt-contrib-watch')
    @loadNpmTasks('grunt-mocha-cli')
    @loadNpmTasks('grunt-mkdir')

    @initConfig
        coffee:
            all:
                options:
                    bare: true
                expand: true,
                cwd: 'src',
                src: ['*.coffee'],
                dest: 'lib',
                ext: '.js'

        clean:
            all: ['lib', 'tmp']

        mkdir:
            all:
                options:
                    create: ['tmp']

        watch:
            all:
                files: ['src/**.coffee', 'test/**.coffee']
                tasks: ['test']

        mochacli:
            options:
                files: 'test/*_test.coffee'
                compilers: ['coffee:coffee-script']
            spec:
                options:
                    reporter: 'spec'

    @registerTask 'npmPack', 'Create NPM package.', ->
        done = @async()

        grunt.util.spawn
            cmd: 'npm'
            args: ['pack']
        , (error, result, code) ->
            grunt.log.writeln(result.stderr) if result.stderr
            grunt.log.writeln(result.stdout) if result.stdout
            done(!error)

    @registerTask 'default', ['test']
    @registerTask 'build', ['clean', 'coffee']
    @registerTask 'package', ['build', 'npmPack']
    @registerTask 'test', ['build', 'mkdir', 'mochacli']
