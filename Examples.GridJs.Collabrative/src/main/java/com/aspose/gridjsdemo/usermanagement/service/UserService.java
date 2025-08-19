package com.aspose.gridjsdemo.usermanagement.service;

import org.hibernate.Hibernate;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.access.prepost.PreAuthorize;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.core.userdetails.UserDetails;
import org.springframework.security.crypto.bcrypt.BCryptPasswordEncoder;
import org.springframework.security.crypto.password.PasswordEncoder;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import com.aspose.gridjsdemo.usermanagement.dto.ChangePasswordForm;
import com.aspose.gridjsdemo.usermanagement.entity.User;
import com.aspose.gridjsdemo.usermanagement.exception.CustomFieldValidationException;
import com.aspose.gridjsdemo.usermanagement.exception.UsernameOrIdNotFound;
import com.aspose.gridjsdemo.usermanagement.repository.UserRepository;

import java.util.Optional;

 
@Service
public class UserService {
    @Autowired
    private UserRepository repository;

    @Autowired
    private PasswordEncoder bCryptPasswordEncoder;

    public Iterable<User> getAllUsers() {
        return repository.findAll();
    }

    private boolean checkUsernameAvailable(User user) throws Exception {
        Optional<User> userFound = repository.findByUsername(user.getUsername());
        if (userFound.isPresent()) {
            throw new CustomFieldValidationException("Username not available", "username");
        }
        return true;
    }

    private boolean checkPasswordValid(User user) throws Exception {
        if (user.getConfirmPassword() == null || user.getConfirmPassword().isEmpty()) {
            throw new CustomFieldValidationException("Confirm Password is required", "confirmPassword");
        }

        if (!user.getPassword().equals(user.getConfirmPassword())) {
            throw new CustomFieldValidationException("Password and Confirm Password are not the same", "password");
        }
        return true;
    }

    public User createUser(User user) throws Exception {
        if (checkUsernameAvailable(user) && checkPasswordValid(user)) {
            String encodedPassword = bCryptPasswordEncoder.encode(user.getPassword());
            user.setPassword(encodedPassword);
            user = repository.save(user);
        }
        return user;
    }
    
    @Transactional 
    public User getUserById(Long id) throws UsernameOrIdNotFound {
    	User u= repository.findById(id).orElseThrow(() -> new UsernameOrIdNotFound("User id does not exist."));
        Hibernate.initialize(u.getRoles()); // 显式初始化 roles 集合
        return u;
    }

    public User updateUser(User fromUser) throws Exception {
        User toUser = getUserById(fromUser.getId());
        mapUser(fromUser, toUser);
        return repository.save(toUser);
    }


    /**
     * Map everything but the password.
     *
     * @param from
     * @param to
     */
    protected void mapUser(User from, User to) {
        to.setUsername(from.getUsername());
        to.setFirstName(from.getFirstName());
        to.setLastName(from.getLastName());
        to.setEmail(from.getEmail());
        to.setRoles(from.getRoles());
    }

    @PreAuthorize("hasAnyRole('ROLE_ADMIN')")
    public void deleteUser(Long id) throws UsernameOrIdNotFound {
        User user = getUserById(id);
        repository.delete(user);
    }

    public User changePassword(ChangePasswordForm form) throws Exception {
        User user = getUserById(form.getId());

        if (!isLoggedUserADMIN() && !user.getPassword().equals(form.getCurrentPassword())) {
            throw new Exception("Current Password invalid.");
        }

        if (user.getPassword().equals(form.getNewPassword())) {
            throw new Exception("New password must be different from the current password.");
        }

        if (!form.getNewPassword().equals(form.getConfirmPassword())) {
            throw new Exception("New Password and Confirm Password do not match.");
        }

        String encodePassword = bCryptPasswordEncoder.encode(form.getNewPassword());
        user.setPassword(encodePassword);
        return repository.save(user);
    }

    private boolean isLoggedUserADMIN() {
        //Get the logged in user
        Object principal = SecurityContextHolder.getContext().getAuthentication().getPrincipal();

        UserDetails loggedUser = null;
        Object roles = null;

        //Verify that this fetched session object is the user
        if (principal instanceof UserDetails) {
            loggedUser = (UserDetails) principal;

            roles = loggedUser.getAuthorities().stream()
                    .filter(x -> "ROLE_ADMIN".equals(x.getAuthority())).findFirst()
                    .orElse(null);
        }
        return roles != null;
    }

    private User getLoggedUser() throws Exception {
        //Get the logged in user
        Object principal = SecurityContextHolder.getContext().getAuthentication().getPrincipal();

        UserDetails loggedUser = null;

        //Verify that this fetched session object is the user
        if (principal instanceof UserDetails) {
            loggedUser = (UserDetails) principal;
        }

        User myUser = repository
                .findByUsername(loggedUser.getUsername()).orElseThrow(() -> new Exception("Error getting the logged " +
                        "in user from the session."));

        return myUser;
    }

}
