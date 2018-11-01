package org.artofsolving.jodconverter.office;

import org.artofsolving.OfficeSoftware;

/**
 * Indicates that the Open/Libre Office software cannot be found.
 * 
 * @author Paul Chapman
 */
public class SoftwareNotFoundException extends RuntimeException {

	private static final long serialVersionUID = -3770098096685690213L;
	private OfficeSoftware preference;

	public SoftwareNotFoundException(String message, OfficeSoftware preference) {
		super(message);
		this.preference = preference;
	}

	public OfficeSoftware getPreference() {
		return preference;
	}

}
