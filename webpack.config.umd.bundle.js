export default {
    mode: 'production',
    entry: {
        xlexport: './src/xlexport.js'
    },
    output: {
        filename: '[name].umd.min.js',
        library: {
            type: 'umd'
        }
    },
    optimization: {
        minimize: true
    }
}
