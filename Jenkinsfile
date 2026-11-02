// Pipeline: Valorizaciones -> Base de datos -> Pedidos SAP
// Jenkins lo lee desde GitHub ("Pipeline script from SCM").
// Funciona en macOS/Linux (sh) y en Windows (bat): detecta el sistema con isUnix().

// Ejecuta Python del entorno virtual según el sistema operativo
def py(String args) {
    if (isUnix()) {
        sh "\"${env.PYTHON_UNIX}\" ${args}"
    } else {
        bat "\"${env.PYTHON_WIN}\" ${args}"
    }
}

pipeline {
    agent any

    environment {
        PYTHON_UNIX = '.venv/bin/python'
        PYTHON_WIN  = '.venv\\Scripts\\python.exe'
        SAP_MODO    = 'simulado'   // "qa" solo con un ambiente de pruebas de SAP
        PYTHONUTF8  = '1'          // evita errores con tildes en la consola de Windows
    }

    triggers {
        cron('H 7 * * 1-5')        // Lunes a viernes, alrededor de las 7 a.m.
        pollSCM('H/5 * * * *')     // Revisa GitHub cada 5 minutos
    }

    stages {
        stage('Preparar Python') {
            steps {
                script {
                    if (isUnix()) {
                        sh 'python3 -m venv .venv'
                    } else {
                        bat 'python -m venv .venv'
                    }
                    py('-m pip install -r requirements.txt')
                }
            }
        }
        stage('Probar lectura de valorizaciones') {
            steps {
                script {
                    py('-m pytest tests')
                    py('src/extraer.py')
                }
            }
        }
        stage('Validación de calidad') {
            steps {
                // Si validar.py termina con exit 1, el pipeline se detiene aquí
                script { py('src/validar.py') }
            }
        }
        stage('Carga a base de datos') {
            steps {
                script { py('src/cargar_bd.py') }
            }
        }
        stage('Generación de pedidos SAP') {
            steps {
                withCredentials([usernamePassword(credentialsId: 'sap-dev',
                                                  usernameVariable: 'SAP_USER',
                                                  passwordVariable: 'SAP_PASS')]) {
                    script { py('src/generar_pedidos.py') }
                }
            }
        }
    }

    post {
        success {
            emailext(
                to: 'michelcabrerayauri998@gmail.com',
                subject: "OK - ${env.JOB_NAME} #${env.BUILD_NUMBER}",
                body: "El pipeline terminó correctamente.\nCommit: ${env.GIT_COMMIT}\nDetalle: ${env.BUILD_URL}"
            )
        }
        failure {
            emailext(
                to: 'michelcabrerayauri998@gmail.com',
                subject: "FALLÓ - ${env.JOB_NAME} #${env.BUILD_NUMBER}",
                body: "El pipeline falló.\nCommit: ${env.GIT_COMMIT}\nRevisa la consola:\n${env.BUILD_URL}console"
            )
        }
    }
}
