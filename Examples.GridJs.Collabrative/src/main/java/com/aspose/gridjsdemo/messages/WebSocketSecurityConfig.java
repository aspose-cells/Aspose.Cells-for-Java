package com.aspose.gridjsdemo.messages;

import org.springframework.context.annotation.Configuration;
import org.springframework.messaging.simp.SimpMessageType;
import org.springframework.security.config.annotation.web.messaging.MessageSecurityMetadataSourceRegistry;
import org.springframework.security.config.annotation.web.socket.AbstractSecurityWebSocketMessageBrokerConfigurer;

@Configuration
public class WebSocketSecurityConfig extends AbstractSecurityWebSocketMessageBrokerConfigurer {

    @Override
    protected void configureInbound(MessageSecurityMetadataSourceRegistry messages) {
        messages
            .simpTypeMatchers(SimpMessageType.CONNECT).authenticated()
            .simpDestMatchers("/app/**","/queue/**").hasRole("USER")   //.authenticated() // 只有 ROLE_USER 用户可以发送消息到 /app/**
            .simpSubscribeDestMatchers("/topic/public").permitAll()
            .simpSubscribeDestMatchers("/user/**", "/topic/private/**", "/topic/messages", "/topic/opr/*").hasRole("USER") //.authenticated() // 只有 ROLE_USER 用户可以发送消息到 /app/**
            .anyMessage().denyAll();
    }

    @Override
    protected boolean sameOriginDisabled() {
        return true; // 开发环境禁用CSRF
    }
}
