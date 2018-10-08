package com.vmware.jct.service.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;

import org.springframework.stereotype.Service;

import com.amazonaws.auth.BasicAWSCredentials;
import com.amazonaws.services.s3.AmazonS3;
import com.amazonaws.services.s3.AmazonS3Client;
import com.amazonaws.services.s3.model.CannedAccessControlList;
import com.amazonaws.services.s3.model.ObjectMetadata;
import com.amazonaws.services.s3.model.PutObjectRequest;
import com.vmware.jct.service.IStorageService;

@Service
public class S3StorageService implements IStorageService {
	
	private static final String BUCKET_NAME = "jobcrafting";
	private static final String AWS_ACCESS_KEY_ID = "";
	private static final String AWS_SECRET_ACCESS_KEY = "";
	
//	static {
//		System.getProperties().put("AWS_ACCESS_KEY_ID", AWS_ACCESS_KEY_ID);
//		System.getProperties().put("AWS_SECRET_ACCESS_KEY", AWS_SECRET_ACCESS_KEY);
//	}

	@Override
	public void store(String key, InputStream in, String contentType) {
		AmazonS3 s3client = new AmazonS3Client(new BasicAWSCredentials(AWS_ACCESS_KEY_ID, AWS_SECRET_ACCESS_KEY));
		ObjectMetadata metadata = new ObjectMetadata();
		metadata.setContentType(contentType);
		PutObjectRequest por = new PutObjectRequest(BUCKET_NAME, key, in, metadata);
		por.setCannedAcl(CannedAccessControlList.PublicRead);
		s3client.putObject(por);
	}
	
	public static void main(String[] args) throws FileNotFoundException {
		String keyName = "images/profile3.jpg";
		String uploadFileName = "/Users/rpunch/Documents/tmp/img1.jpg";
		S3StorageService s3 = new S3StorageService();
		s3.store(keyName, new FileInputStream(new File(uploadFileName)), "image/jpeg");
	}

}
