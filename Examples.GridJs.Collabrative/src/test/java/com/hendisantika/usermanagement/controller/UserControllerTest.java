package com.hendisantika.usermanagement.controller;

import com.aspose.gridjsdemo.usermanagement.controller.UserController;
import com.aspose.gridjsdemo.usermanagement.dto.ChangePasswordForm;
import com.aspose.gridjsdemo.usermanagement.entity.Role;
import com.aspose.gridjsdemo.usermanagement.entity.User;
import com.aspose.gridjsdemo.usermanagement.repository.RoleRepository;
import com.aspose.gridjsdemo.usermanagement.service.UserService;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;
import org.springframework.http.MediaType;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.setup.MockMvcBuilders;

import java.util.Collections;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyLong;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.doNothing;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultHandlers.print;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.redirectedUrl;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

@ExtendWith(MockitoExtension.class)
class UserControllerTest {

    private MockMvc mockMvc;

    @InjectMocks
    private UserController sut;

    @Mock
    private UserService userService;

    @Mock
    private RoleRepository roleRepository;


    private static User user;
    private static Role role1;

    @BeforeAll
    public static void setUpBeforeAll() {
        user = new User();
        user.setId(1L);
        user.setFirstName("Piolo");
        user.setLastName("Pascual");
        user.setEmail("a@a.com");
        user.setUsername("ppascual");
        user.setPassword("dsa");

        role1 = new Role();
        role1.setId(1L);
        role1.setName("SUPER ADMIN");
        role1.setDescription("ROLE SUPER ADMIN");
    }

    @BeforeEach
    public void setUpBefore() {
        mockMvc = MockMvcBuilders.standaloneSetup(sut).build();
    }

    @Test
    void testCreateUser() throws Exception {
        when(userService.createUser(any(User.class))).thenReturn(user);

        mockMvc.perform(post("/userForm")
                        .flashAttr("userForm", user))
                .andDo(print())
                .andExpect(status().isOk());
    }

    @Test
    void testGetEditUserForm() throws Exception {
        when(userService.getUserById(anyLong())).thenReturn(user);

        mockMvc.perform(get("/editUser/{id}", 1L))
                .andDo(print())
                .andExpect(status().isOk());
    }

    @Test
    void testLoginAndReturnToIndex() throws Exception {
        mockMvc.perform(get("/"))
                .andDo(print())
                .andExpect(status().isOk());

        mockMvc.perform(get("/login"))
                .andDo(print())
                .andExpect(status().isOk());
    }

    @Test
    void testSignUp() throws Exception {
        Role role1 = new Role();
        role1.setId(1L);
        role1.setName("SUPER ADMIN");
        role1.setDescription("ROLE SUPER ADMIN");
        when(roleRepository.findAll()).thenReturn(Collections.emptyList());
        when(roleRepository.save(any(Role.class))).thenReturn(role1);
        when(roleRepository.findByName(anyString())).thenReturn(role1);

        mockMvc.perform(get("/signup"))
                .andDo(print())
                .andExpect(status().isOk());
    }

    @Test
    void testSignUpAction() throws Exception {
        when(roleRepository.findByName(anyString())).thenReturn(role1);
        when(userService.createUser(any(User.class))).thenReturn(user);

        mockMvc.perform(post("/signup")
                        .flashAttr("userForm", user))
                .andDo(print())
                .andExpect(status().isOk());
    }
/*
    @Test
    void testUserForm() throws Exception {
        mockMvc.perform(get("/userForm"))
                .andDo(print())
                .andExpect(status().isOk());
    }
    */

    @Test
    void testPostEditUserForm() throws Exception {
        when(userService.updateUser(any(User.class))).thenReturn(user);

        mockMvc.perform(post("/editUser")
                        .flashAttr("userForm", user))
                .andDo(print())
                .andExpect(status().isOk());
    }

    @Test
    void testCancelEditUser() throws Exception {
        mockMvc.perform(get("/userForm/cancel"))
                .andDo(print())
                .andExpect(status().isFound())
                .andExpect(redirectedUrl("/userForm"));
    }
    /*
    @Test
    void testDeleteUser() throws Exception {
        doNothing().when(userService).deleteUser(anyLong());

        mockMvc.perform(get("/deleteUser/{id}", 1L))
                .andDo(print())
                .andExpect(status().isOk());
    }
    */

    @Test
    void testPostEditUseChangePassword() throws Exception {

        ChangePasswordForm form = new ChangePasswordForm();
        form.setId(1L);
        form.setCurrentPassword("dasd");
        form.setNewPassword("awsd");
        form.setConfirmPassword("awsd");

        when(userService.changePassword(any(ChangePasswordForm.class))).thenReturn(user);

        mockMvc.perform(post("/editUser/changePassword")
                        .contentType(MediaType.APPLICATION_JSON)
                        .accept(MediaType.APPLICATION_JSON)
                        .content(new ObjectMapper().writeValueAsString(form)))
                .andDo(print())
                .andExpect(status().isOk());
    }

}
