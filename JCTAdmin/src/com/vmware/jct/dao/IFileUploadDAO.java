package com.vmware.jct.dao;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctOnetOccupationList;

public interface IFileUploadDAO {
	/**
	 * Method saves new OnetOccupation data
	 * @param jctOnetOccupation
	 * @throws DAOException
	 */
	public String saveONetData(JctOnetOccupationList jctOnetOccupation) throws DAOException;
	/**
	 * Method to fetch the occupation title against the occupation code
	 * @param occupationCode
	 * @throws DAOException
	 */
	public String getOccupationTitle(String occupationCode) throws DAOException;
	/**
	 * Method to update the occupation data
	 * @param jctOnetOccupation
	 * @throws DAOException
	 */
	public String updateONetData(JctOnetOccupationList jctOnetOccupation) throws DAOException;
	/**
	 * Method to fetch the occupation object against the occupation code
	 * @param occupationCode
	 * @throws DAOException
	 */
	public JctOnetOccupationList getOccupationObj(String occupationCode) throws DAOException;
}
