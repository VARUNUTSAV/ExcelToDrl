pipeline {
    agent any
    stages {
        stage('Clone Repository') {
            steps {
                git url: 'https://bitbucket.org/your-repo.git', branch: 'master'
            }
        }
        stage('Run Excel to DRL Converter') {
            steps {
                sh 'mvn compile exec:java -Dexec.args="src/main/resources"'
            }
        }
    }
}