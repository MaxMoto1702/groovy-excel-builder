import org.asciidoctor.gradle.jvm.AsciidoctorTask
import org.gradle.api.publish.maven.MavenPublication

plugins {
    id("org.asciidoctor.jvm.convert") version "4.0.5"
    kotlin("jvm") version "2.2.0"
    id("io.gitlab.arturbosch.detekt") version "1.23.8"
    id("org.jlleitschuh.gradle.ktlint") version "12.1.2"
    id("maven-publish")
}

group = "maxmoto1702"
version = "1.1.0"

repositories {
    mavenCentral()
}

dependencies {
    implementation("org.apache.poi:poi:5.5.1")
    implementation("org.apache.poi:poi-ooxml:5.5.1")
    implementation("org.slf4j:slf4j-api:2.0.17")

    testImplementation("org.junit.jupiter:junit-jupiter:5.13.4")
    testImplementation("org.assertj:assertj-core:3.27.7")
    testImplementation("org.apache.logging.log4j:log4j-api:2.25.3")
    testImplementation("org.apache.logging.log4j:log4j-core:2.25.3")
    testImplementation("org.apache.logging.log4j:log4j-slf4j2-impl:2.25.3")
    testImplementation("org.slf4j:jcl-over-slf4j:2.0.17")
    testRuntimeOnly("org.junit.platform:junit-platform-launcher")

    detektPlugins("io.gitlab.arturbosch.detekt:detekt-formatting:1.23.8")
}

kotlin {
    jvmToolchain(21)
}

detekt {
    buildUponDefaultConfig = true
    allRules = false
    parallel = true
    config.setFrom(files("$rootDir/config/detekt/detekt.yml"))
}

ktlint {
    android.set(false)
    ignoreFailures.set(false)
    outputToConsole.set(true)
}

tasks.test {
    useJUnitPlatform()
}

tasks.withType<AsciidoctorTask>().configureEach {
    logDocuments = true
    outputOptions {
        backends("html5")
    }
    attributes(
        mapOf(
            "toc" to "left",
            "sourceDir" to "$projectDir/src/main/kotlin",
            "version" to project.version,
            "groupId" to project.group,
            "artifactId" to project.name,
        ),
    )
}

publishing {
    publications {
        create<MavenPublication>("mavenJava") {
            from(components["java"])
        }
    }
    repositories {
        maven {
            url = uri(layout.buildDirectory.dir("repo"))
        }
    }
}
