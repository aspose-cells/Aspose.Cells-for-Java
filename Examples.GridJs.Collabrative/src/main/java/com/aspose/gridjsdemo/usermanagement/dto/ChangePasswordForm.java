package com.aspose.gridjsdemo.usermanagement.dto;

//import jakarta.validation.constraints.NotBlank;
//import jakarta.validation.constraints.NotNull;
import javax.validation.constraints.NotBlank;
import javax.validation.constraints.NotNull;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;

 
@Data
@AllArgsConstructor
@NoArgsConstructor
@EqualsAndHashCode
public class ChangePasswordForm {
    @NotNull
    private Long id;

    @NotBlank(message = "Current Password must not be blank")
    private String currentPassword;

    @NotBlank(message = "New Password must not be blank")
    private String newPassword;

    @NotBlank(message = "Confirm Password must not be blank")
    private String confirmPassword;

    public ChangePasswordForm(Long id) {
        this.id = id;
    }
}
