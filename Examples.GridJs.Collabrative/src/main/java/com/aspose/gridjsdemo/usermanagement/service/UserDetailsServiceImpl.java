package com.aspose.gridjsdemo.usermanagement.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.authority.SimpleGrantedAuthority;
import org.springframework.security.core.userdetails.User;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.security.core.userdetails.UsernameNotFoundException;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.aspose.gridjsdemo.usermanagement.dto.CustomUserDetails;
import com.aspose.gridjsdemo.usermanagement.entity.Role;
import com.aspose.gridjsdemo.usermanagement.repository.UserRepository;

import java.util.HashSet;
import java.util.Set;
 
@Service
@Transactional
public class UserDetailsServiceImpl implements UserDetailsService {
    @Autowired
    private UserRepository userRepository;

    @Override
    public CustomUserDetails loadUserByUsername(String username) throws UsernameNotFoundException {

        com.aspose.gridjsdemo.usermanagement.entity.User appUser =
                userRepository.findByUsername(username).orElseThrow(() -> new UsernameNotFoundException("Login " +
                        "Username Invalid."));

        Set<GrantedAuthority> grantList = new HashSet<GrantedAuthority>();
        for (Role role : appUser.getRoles()) {
            GrantedAuthority grantedAuthority = new SimpleGrantedAuthority(role.getDescription());
            grantList.add(grantedAuthority);
        }

        return new CustomUserDetails(appUser.getId(),appUser.getUsername(), appUser.getPassword(), grantList);
    }
}
