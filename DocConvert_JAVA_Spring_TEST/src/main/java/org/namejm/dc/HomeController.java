package org.namejm.dc;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Map;

import DocConvert_API.DocConvert;
import DocConvert_API.WebCapture;
import DocConvert_API.getProperties;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.view.AbstractView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * Handles requests for the application home page.
 */
@Controller
public class HomeController{
	private static final Logger logger = LoggerFactory.getLogger(HomeController.class);
	private static String SAVE_PATH = "";

	@RequestMapping(value = "/", method = RequestMethod.GET)
	public String home(HttpServletRequest request) {
		HttpServletRequest req = ((ServletRequestAttributes) RequestContextHolder.currentRequestAttributes()).getRequest();
		logger.info("▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶");
		logger.info("▶▶▶▶▶ IP: " + req.getRemoteAddr());
		logger.info("▶▶▶▶▶ CLASS: " + HomeController.class);
		logger.info("▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶▶");
		dcProperties dcProperties = new dcProperties();
		dcProperties.readProperties();
		SAVE_PATH = dcProperties.workspacePath();

		return "home";
	}

	@RequestMapping(value = "/fileUpload", method = RequestMethod.POST)
	public ModelAndView upload(@RequestParam(value="file1", required = false) MultipartFile mf) throws Exception {
		String originalFileName = mf.getOriginalFilename();
		long fileSize = mf.getSize();
		String saveFile = SAVE_PATH + File.separator + System.currentTimeMillis() + originalFileName;
		try {
			mf.transferTo(new File(saveFile));
		} catch (IllegalStateException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		File fileAbsolutePath = new File(saveFile);
		String isSuccess = new DocConvert().DocConvert_Start(fileAbsolutePath.getParent(), fileAbsolutePath.getParent(), fileAbsolutePath.getName(), 2);
		new File(saveFile).delete();

		String sourceFileExten = saveFile.substring(saveFile.lastIndexOf("."), saveFile.length());
		File downloadFile = new File(saveFile.replace(sourceFileExten, ".zip"));

		if(!downloadFile.exists())
			return null;

		return new ModelAndView("download", "downloadFile", downloadFile);
	}

	@RequestMapping(value = "/webCapture", method = RequestMethod.POST)
	public ModelAndView webCapture(@RequestParam("URL") String url) throws Exception {

		String saveFile = SAVE_PATH + File.separator + System.currentTimeMillis();
		File downloadFile = new File(saveFile + ".png");
		String isSuccess = new WebCapture().WebCapture_Start(url, downloadFile.getAbsolutePath());

		if(!downloadFile.exists())
			return null;

		return new ModelAndView("download", "downloadFile", downloadFile);
	}
}
