node {
      def app
      environment {
        currentDateTime = sh(returnStdout: true, script: 'date -d \'+3 hour\' +%d%m%Y%H%M%S').trim()
      }

      stage('Clone repository') {

            checkout scm
      }
      stage("Docker build"){
        app = docker.build("migutak/docxv2")
      }

      stage('Test'){

        script {
          DATE_TAG = java.time.LocalDate.now()
          DATETIME_TAG = java.time.LocalDateTime.now()
        }
        sh "echo ${DATETIME_TAG}"

        }

      stage('Push image') {
        /* Finally, we'll push the image with two tags:
         * First, the incremental build number from Jenkins
         * Second, the 'latest' tag.
         * Pushing multiple tags is cheap, as all the layers are reused. */
        docker.withRegistry('https://registry.hub.docker.com', 'docker_credentials') {
            app.push("${env.BUILD_NUMBER}.${currentDateTime}")
            app.push("latest")
        }
      }
    }
