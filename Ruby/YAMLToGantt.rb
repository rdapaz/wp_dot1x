# coding: utf-8
#!/usr/bin/ruby
%w( yaml pp win32ole ).each { |dep| require dep }

class WIN32OLE
  def self.get_project
    begin
      connect('MSProject.Application')
    rescue WIN32OLERuntimeError => e
      new('MSProject.Application')
    end
  end
end

class MicrosoftProject
  attr_accessor :doc_file

  def initialize(the_doc)
    @doc_file = the_doc
    @app = WIN32OLE::get_project
    @app.Visible = true
    @doc = @app.FileOpen(@doc_file)
    @proj = @app.ActiveProject
    @task_ids = []
  end

  def add_new_task(args)
    task_name = args[:task_name]
    durn = args[:durn]
    resources = args[:resources]
    nesting = args[:nesting] + 1
    tsk = @proj.Tasks.Add(Name: task_name)
    tsk.Duration = durn if durn
    tsk.Text1 = nesting
    tsk.ResourceNames = resources if resources
    unless @task_ids.size == 0 || @task_ids[-1][1] != nesting
      tsk.Predecessors = @task_ids[-1][0]
    end
    tsk.OutlineIndent while tsk.OutlineLevel < nesting
    tsk.OutlineOutdent while tsk.OutlineLevel > nesting
    @task_ids << [tsk.ID, nesting]
  end

  def yaml_to_gantt(obj, nesting=0)
    counter = 0
    obj.each do | task, rest |
      counter += 1
      if rest.is_a? Hash
        add_new_task(  :task_name => task,
                       :nesting => nesting)
        yaml_to_gantt(rest, nesting+1)
      else
        durn, resources = rest.split(/\|/,2)
        add_new_task(  :task_name => task,
                       :durn => durn,
                       :resources => resources,
                       :nesting => nesting)
      end
    end
  end
end

def main(obj)
  filename = 'D:\Projects\Western Power\NEW\temp.mpp' # Requires full path for some reason
  pj = MicrosoftProject.new(filename)
  pj.yaml_to_gantt(obj)
end


def test_me(o, nesting=0)
  spaces = %q{  }
  o.each do |task, rest|
    if rest.is_a? Hash
      puts spaces * nesting + task
      test_me(rest, nesting+1)
    else
      puts [spaces * nesting + task, rest].join(": ")
    end
  end
  1
end

if __FILE__ == $0
  SOURCE = 'D:\__NEW__\scripts\tasks.yaml'
  obj = File.exists?(SOURCE) ? YAML.load_file(SOURCE) : YAML::load(DATA)
  test_me(obj)
  main(obj)
end

__END__
---
<<Project Name>>:
    Initiation:
        Assessment Workshop:
            MRA Completed: 0.0d|Principle
            Risk Assessment Completed: 0.0d|
            OCM Assessment Completed: 0.0d|
            Commercial Assessment Completed: 1.0d|
            Documents Assessment Completed: 0.0d|