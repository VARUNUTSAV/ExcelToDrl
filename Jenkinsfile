pipeline {
    agent any
    stages {
        stage('Clone Repository') {
            steps {
                git url: 'https://github.com/VARUNUTSAV/ExcelToDrl', branch: 'main'
            }
        }
        stage('Run Excel to DRL Converter') {
            steps {
                sh 'mvn compile exec:java -Dexec.args="src/main/resources"'
            }
        }
    }
}
