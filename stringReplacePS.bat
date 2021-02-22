SET app_nome="app_nome"

SET valve_kc="^<Valve className=^"org.keycloak.adapters.tomcat.KeycloakAuthenticatorValve^"/^>"
SET valve_j="^<!--Valve className=^"org.keycloak.adapters.tomcat.KeycloakAuthenticatorValve^"/--^>"

SET login_config_form_auth_method_kc="^<auth-method>BASIC^</auth-method^>"
SET login_config_form_realm_name_kc="^<realm-name^>TRT15^</realm-name^>"

SET login_config_form_auth_method_j="^<auth-method^>FORM^</auth-method^>"
SET login_config_form_realm_name_j="^<realm-name^>default^</realm-name^>^<form-login-config^>^<form-login-page^>/res-plc/login/loginPlc.html^</form-login-page^>^<form-error-page^>/res-plc/login/loginErroPlc.html^</form-error-page^>^</form-login-config^>"

SET security_role_open_kc="^<security-role^>"
SET security_role_open_j="^<!--security-role--^>"

SET security_role_close_kc="^</security-role^>"
SET security_role_open_j="^<!--/security-role--^>"

REM META-INF/context.xml
powershell -Command "(gc webapps/%app_nome%/META-INF/context.xml) -replace %valve_kc%, %valve_j% | Out-File -encoding ASCII webapps/%app_nome%/META-INF/context.xml"

REM WEB-INF/web.xml
powershell -Command "(gc webapps/%app_nome%/WEB-INF/web.xml) -replace '^<role-name>uma_authorization</role-name>', '^<!--role-name>uma_authorization</role-name-->' | Out-File -encoding ASCII webapps/%app_nome%/WEB-INF/web.xml"
powershell -Command "(gc webapps/%app_nome%/WEB-INF/web.xml) -replace %login_config_form_auth_method_kc%, %login_config_form_auth_method_j% | Out-File -encoding ASCII webapps/%app_nome%/WEB-INF/web.xml"
powershell -Command "(gc webapps/%app_nome%/WEB-INF/web.xml) -replace %security_role_open_kc%, %security_role_open_j% | Out-File -encoding ASCII webapps/%app_nome%/WEB-INF/web.xml"
powershell -Command "(gc webapps/%app_nome%/WEB-INF/web.xml) -replace %security_role_close_kc%, %security_role_open_j% | Out-File -encoding ASCII webapps/%app_nome%/WEB-INF/web.xml"

REM conf/.../context.xml
powershell -Command "(gc conf/Catalina/localhost/%app_nome%.xml) -replace %valve_kc%, %valve_j% | Out-File -encoding ASCII conf/Catalina/localhost/%app_nome%.xml"

pause
