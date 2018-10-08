package com.vmware.jct.common.utility;
import java.util.Comparator;

public class PetComparator implements Comparator {

	int ocuupationId;
	public String occupationTitle;
	

	public PetComparator(int ocuupationId, String occupationTitle){
		this.ocuupationId= ocuupationId;
		this.occupationTitle = occupationTitle;
	}
	
	public PetComparator() {
		// TODO Auto-generated constructor stub
	}

	@Override
	public int compare(Object o1, Object o2) {
		
		int result = 0;
		
		PetComparator pet = (PetComparator)o1;
		PetComparator petOther = (PetComparator)o2;
		
		Integer ocuupationIdInt = new Integer(pet.ocuupationId);
		Integer ocuupationIdOtherInt = new Integer(petOther.ocuupationId);
		
		result = ocuupationIdInt.compareTo(ocuupationIdOtherInt);
		
		if(result==0){
			result = pet.occupationTitle.compareTo(petOther.occupationTitle);
		}
		
		return result;
	}
	
	
	@Override
	public String toString() {
		
		return occupationTitle;
	}

	
}
