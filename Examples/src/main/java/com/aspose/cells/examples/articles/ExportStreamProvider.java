package com.aspose.cells.examples.articles;

import java.io.File;
import java.io.FileOutputStream;

import com.aspose.cells.IStreamProvider;
import com.aspose.cells.StreamProviderOptions;



public class ExportStreamProvider implements IStreamProvider {
	private String outputDir;

	public ExportStreamProvider(String dir) {
		outputDir = dir;
		System.out.println(outputDir);
	}

	@Override
	public void closeStream(StreamProviderOptions options) throws Exception {
		if (options != null && options.getStream() != null) {
			options.getStream().close();
		}
	}

	@Override
	public void initStream(StreamProviderOptions options) throws Exception {
		System.out.println(options.getDefaultPath());

		File file = new File(outputDir);
		if (!file.exists() && !file.isDirectory()) {
			file.mkdirs();
		}
		String defaultPath = options.getDefaultPath();
		String path = outputDir + defaultPath.substring(defaultPath.lastIndexOf("/") + 1);
		options.setCustomPath(path);
		options.setStream(new FileOutputStream(path));
	}
}

