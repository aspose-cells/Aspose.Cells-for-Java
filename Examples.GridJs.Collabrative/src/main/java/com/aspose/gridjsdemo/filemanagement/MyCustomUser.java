package com.aspose.gridjsdemo.filemanagement;

import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;

import com.aspose.gridjs.CoWorkUserPermission;
import com.aspose.gridjs.CoWorkUserProvider;
import com.aspose.gridjsdemo.usermanagement.dto.CustomUserDetails;

public class MyCustomUser implements CoWorkUserProvider{

	@Override
	public String getCurrentUserName() {
		 
		return getCurrentUser().getUsername();
	}

	@Override
	public Long getCurrentUserId() {
		 
		return getCurrentUser().getUserId();
	}

	@Override
	public CoWorkUserPermission getPermission() {
	 
		return CoWorkUserPermission.EDITABLE;
	}
	
	 public static CustomUserDetails getCurrentUser() {
	        Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
	        if (authentication != null && authentication.getPrincipal() instanceof CustomUserDetails) {
	            return (CustomUserDetails) authentication.getPrincipal();
	        }
	        throw new IllegalStateException("User not authenticated");
	    }

}
