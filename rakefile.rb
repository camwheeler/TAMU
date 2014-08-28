require 'albacore'
require 'version_bumper'

basePath = File.expand_path(File.dirname(__FILE__)) 
zipDir = "#{basePath}/zip"
buildDir = "#{basePath}/build"
release = "#{buildDir}/release"
assemblyInfoDir = "#{basePath}/src/TimeAndMetricsUpdater/Properties"
v = "#{bumper_version.to_s}.#{ENV['BUILD_NUMBER'] || '9999'}"

desc "Default"
task :default => [:clean, :build, :merge, :zip]
task :zip => [:moveStuff, :createZip]

desc "Clean out previous build artifacts"
task :clean do
   FileUtils.mkdir(assemblyInfoDir) unless File.exists?(assemblyInfoDir)
   FileUtils.mkdir(buildDir) unless File.exists?(buildDir)
   FileUtils.mkdir(zipDir) unless File.exists?(zipDir)
   FileUtils.rm_rf(Dir.glob("#{buildDir}/*"))
   FileUtils.rm_rf(Dir.glob("#{zipDir}/*"))
   FileUtils.mkdir(release)
end

desc "Build from source"
msbuild :build => [:solutionNugetPackages, :assemblyinfo] do |msb|
   msb.properties = { 
      :configuration => :Release,
      :Platform => "x64", 
      :outdir => buildDir}
   msb.targets = [ :Clean, :Build ]
   msb.solution = "#{basePath}/src/TimeAndMetricsUpdater.sln"
end

desc "Merge built assemblies"
ilmerge :merge do |cfg|
   puts "Merging assemblies..."
   assemblies = Dir.glob("#{buildDir}/*.dll")
   cfg.command = "ILMerge.exe"
   cfg.assemblies "/targetplatform:v4 #{buildDir}/TimeAndMetricsUpdater.exe", assemblies
   cfg.output = "#{release}/TAMU.exe"
end

desc "Restore nuget packages."
exec :solutionNugetPackages do |cmd|
   cmd.command = "nuget.exe"
   cmd.parameters "install src/TimeAndMetricsUpdater/packages.config -OutputDirectory src/packages"
end

desc "Create and add version to AssemblyInfo"
assemblyinfo :assemblyinfo do |asm|
   asm.output_file = "#{assemblyInfoDir}/AssemblyInfo.cs"
   asm.version = asm.file_version = v
   asm.product_name = "Time and Metrics Updater"  
   asm.company_name = "Zeroworks"
end

desc "Move files into directory that will get zipped"
task :moveStuff do
   FileUtils.cp("#{buildDir}/TimeAndMetricsUpdater.exe.config", "#{release}/TAMU.exe.config")
   FileUtils.cp("#{buildDir}/alarm.ico", "#{release}/alarm.ico")
   FileUtils.cp("settings/oauth.parameters", "#{release}/oauth.parameters")
end

zip :createZip do | zip |
   zip.directories_to_zip release
   zip.output_file = "TimeAndMetricsUpdater.#{v}.zip"
   zip.output_path = zipDir
end