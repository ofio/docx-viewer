import typescript from '@rollup/plugin-typescript';
import terser from '@rollup/plugin-terser';
import nodeExternals from 'rollup-plugin-node-externals';

const umdOutput = {
	name: "docx",
	file: 'docs/js/docx-preview.js',
	sourcemap: true,
	format: 'umd',
	globals: {
		jszip: 'JSZip',
		konva: 'Konva',
		"lodash-es": '_',
	}
};

export default args => {
	const config = {
		input: 'src/docx-preview.ts',
		output: [umdOutput],
		plugins: [
			nodeExternals(),
			typescript(),
		],
	}

	if (args.environment === 'BUILD:production') {
		// 输出配置
		config.output = [umdOutput,
			{
				...umdOutput,
				file: 'dist/docx-preview.js',
			},
			{
				...umdOutput,
				file: 'dist/docx-preview.min.js',
				plugins: [terser()]
			},
			{
				file: 'dist/docx-preview.esm.js',
				sourcemap: true,
				format: 'es',
			},
			{
				file: 'dist/docx-preview.esm.min.js',
				sourcemap: true,
				format: 'es',
				plugins: [terser()]
			},
		];
	}

	return config
};
