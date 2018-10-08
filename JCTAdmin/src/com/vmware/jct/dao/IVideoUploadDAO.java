package com.vmware.jct.dao;

import java.util.List;

import com.vmware.jct.exception.DAOException;
import com.vmware.jct.model.JctPopupInstruction;

public interface IVideoUploadDAO {
	/**
	 * Method fetches videos based on profile and page
	 * @param profileId
	 * @param insPage
	 * @return
	 * @throws DAOException
	 */
	public List<JctPopupInstruction> getElements (int profileId, String insPage) throws DAOException;
	/**
	 * Method updates instruction
	 * @param obj
	 * @return
	 * @throws DAOException
	 */
	public String updateEntity (JctPopupInstruction obj) throws DAOException;
	/**
	 * Method saves instruction
	 * @param obj
	 * @return
	 * @throws DAOException
	 */
	public String saveEntity (JctPopupInstruction obj) throws DAOException;
	/**
	 * Method fetches videos based on selection
	 * @param profileId
	 * @param page
	 * @return
	 * @throws DAOException
	 */
	public List<JctPopupInstruction> getSelectedElements (int profileId, String page) throws DAOException;
	/**
	 * Method fetches instruction based on selection
	 * @param profileId
	 * @param page
	 * @return
	 * @throws DAOException
	 */
	public List<JctPopupInstruction> getSelectedInstruction(int profileId,String page) throws DAOException;
}
