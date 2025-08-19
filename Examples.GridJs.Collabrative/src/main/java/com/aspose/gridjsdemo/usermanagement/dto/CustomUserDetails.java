package com.aspose.gridjsdemo.usermanagement.dto;

import java.util.Collection;

import org.springframework.security.core.GrantedAuthority;
import org.springframework.security.core.userdetails.UserDetails;

public class CustomUserDetails implements UserDetails {
  
	private static final long serialVersionUID = 1L;
	private final Long userId;  // 业务系统唯一ID（不可变）
    private final String username; // 登录用户名（可修改）
    private final String password;
    private final Collection<? extends GrantedAuthority> authorities;
    
    public CustomUserDetails(Long userId, String username, String password, 
                           Collection<? extends GrantedAuthority> authorities) {
        this.userId = userId;
        this.username = username;
        this.password = password;
        this.authorities = authorities;
    }
    
    // 必须实现的方法
    @Override public String getUsername() { return username; }
    @Override public String getPassword() { return password; }
    @Override public Collection<? extends GrantedAuthority> getAuthorities() { return authorities; }
    
    // 自定义扩展方法
    public Long getUserId() { return userId; }
    
    
	 

    // 其他UserDetails默认实现（账户状态检查）
    @Override public boolean isAccountNonExpired() { return true; }
    @Override public boolean isAccountNonLocked() { return true; }
    @Override public boolean isCredentialsNonExpired() { return true; }
    @Override public boolean isEnabled() { return true; }
}
