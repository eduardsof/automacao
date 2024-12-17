pipeline {
    agent any
    stages {
        stage('Clone Repository') {
            steps {
                echo "Clonando repositório..."
                checkout scm
            }
        }
        stage('Run Script') {
            steps {
                echo "Executando o teste.py..."
                bat 'python teste.py' // Para Windows
            }
        }
    }
}
