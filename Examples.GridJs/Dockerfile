
 
# FROM mcr.microsoft.com/java/openjdk:11  AS build
  FROM eclipse/ubuntu_jdk8

# 设置工作目录
WORKDIR /usr/src/app

 


# 将本地的 Maven 设置文件和环境变量设置文件复制到容器中
COPY mvnw .
COPY .mvn .mvn
COPY pom.xml .
COPY src src

# 使用 Maven 构建应用（假设你使用的是 Maven）
RUN ./mvnw package -DskipTests





# Set the user
USER root

RUN ls -l *

# 将构建的 jar 文件复制到新的容器中
COPY  target/*.jar /app/app.jar

# 删除不再需要的源码和 Maven 缓存，减少镜像体积
RUN rm -rf target && rm -rf .mvn && rm -rf src




 



# web port
EXPOSE 8080
# Install Fonts because the SDK image contains very few fonts. The command copies font files from local to docker image. Make sure you have a local “fonts” directory that contains all the fonts you need to install. In this example, the local “fonts” directory is put in the same directory as the Dockerfile.
# COPY fonts/* /usr/share/fonts/
RUN mkdir -p /app/wb
RUN mkdir -p /app/grid_cache/streamcache
COPY wb/*.xlsx /app/wb/
# RUN ls -l /app/
# 启动应用
ENTRYPOINT ["java","-Djava.security.egd=file:/dev/./urandom","-jar","/app/app.jar"]