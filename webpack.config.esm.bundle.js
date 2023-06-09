export default {
    mode: 'production',
    experiments: {
        outputModule: true
    },
    entry: {
        xlexport: './src/xlexport.js'
    },
    output: {
        filename: '[name].esm.min.js',
        library: {
            type: 'module'
        }
    },
    optimization: {
        minimize: true
    }
}
