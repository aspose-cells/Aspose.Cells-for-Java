package com.aspose.gridjsdemo.usermanagement.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.security.authentication.dao.DaoAuthenticationProvider;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
import org.springframework.security.core.userdetails.UserDetailsService;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.security.web.SecurityFilterChain;
import org.springframework.security.web.util.matcher.AntPathRequestMatcher;

import lombok.RequiredArgsConstructor;

 
@Configuration
@EnableWebSecurity
@RequiredArgsConstructor
public class WebSecurityConfig {
    private final String[] PUBLIC_LINK = new String[]{
            "/include/**", "/css/**","/jslib/**", "/icons/**", "/img/**", "/js/**", "/layer/**", "/static/**"
    };

    private final PasswordEncoder bCryptPasswordEncoder;

    private final UserDetailsService userDetailsService;

    @Bean
    public DaoAuthenticationProvider authenticationProvider() {
        DaoAuthenticationProvider auth = new DaoAuthenticationProvider();
        auth.setUserDetailsService(userDetailsService);
        auth.setPasswordEncoder(bCryptPasswordEncoder);
        return auth;
    }

   
//    public void addResourceHandlers(ResourceHandlerRegistry registry) {
//        registry
//            .addResourceHandler("/static/**")
//            .addResourceLocations("classpath:/static/")
//            .setCachePeriod(0); // 开发时禁用缓存
//    }
    
    @Bean
    public SecurityFilterChain configure(HttpSecurity http) throws Exception {
    
//		http.csrf(AbstractHttpConfigurer::disable)
		http.csrf(csrf -> csrf
                .ignoringAntMatchers("/ws/**").ignoringAntMatchers("/GridJs2/**"))  // 特别排除WebSocket端点
				.authorizeHttpRequests(req -> req.antMatchers(PUBLIC_LINK).permitAll()
						.antMatchers("/ws/**").permitAll() //   允许访问 WebSocket 端点
												.antMatchers("/gridjsdemo/**", "/GridJs2/**").authenticated()
												.antMatchers("/", "/index", "/signup", "/test-load", "/raw-file", "/deep-check").permitAll()
												.anyRequest().authenticated())
				.formLogin(formLogin -> formLogin.loginPage("/login").permitAll()
							.defaultSuccessUrl("/userForm")
							.failureUrl("/login?error=true")
							.usernameParameter("username")
							.passwordParameter("password")

				).logout(logout -> logout.logoutRequestMatcher(new AntPathRequestMatcher("/logout"))
						.logoutSuccessUrl("/login?logout").permitAll());

        return http.build();
    }
    
    
    
    
    
 
}
