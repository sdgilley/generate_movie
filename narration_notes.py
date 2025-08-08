# Narration Notes for PowerPoint Slides
# 
# Format: Each entry should match the exact title of your PowerPoint slide
# The key is the slide title, and the value is the narration text
#
# Example:
# "Slide Title Here": "This is the narration that will be spoken for this slide."

narration_notes = {
    "Code Snippet Maintenance in Documentation": "Welcome. This short video introduces the workflow for code maintenance in documentation, specifically what we do for Azure Machine Learning and Azure AI Foundry.",
    
    "Definitions": "First a few definitions. There are three different types of repositories involved in this process. To try to keep them straight, we'll call them the Code repo, the Docs repo, and the Maintenance repo.",
    
    "Learn site": "Here's an example article on the learn site, which contains some code.",
    
    "GitHub": "Here's the actual code content, stored in a Code Repo. Note the comments that start and stop the block of code called chat_completion.",
    
    "How to reference": "To use code from an external Code Repo, set a path_to_root entry in the config file. I use the repo name followed by the branch name.",
    
    "Article markdown": "Now that you have a path_to_root, use it to reference the file. Use id to get to the particular block you want to display.",
    
    "Advantages": "There are many advantages. The code is authored in an editor that can show typos or mistakes. It's runnable. Styling can be enforced. The repo is maintained by teams using the code. It can be set up with automated testing. And it provides a single source of truth.",
    
    "PROBLEMS": "Problems include file deletion, renaming, or content changes that break references. Also, updates to code files aren't reflected in docs until rebuilt.",
    
    "SOLUTION: Monitor the Code Repo": "We monitor the code repo using CODEOWNERS to protect referenced files. Any changes require our review.",
    
    "Maintenance Process": "Our process includes daily and weekly tasks using Python scripts in the Maintenance Repo. These run in Codespaces and rotate monthly among team members.",
    
    "Daily: 1. run find-prs": "We check for PRs needing review across multiple repos using a script that outputs a markdown report.",
    
    "2. run pr-report": "Open the markdown file and view the table. Run the code to check for issues before approving PRs.",
    
    "3. approve PR": "Approve the PR if no issues. If problems exist, fix the docs first. Consider release branches during events like Build or Ignite.",
    
    "Weekly: 1. find-snippets": "Search docs for changes in referenced files. This builds CODEOWNERS files and a CSV for lookup.",
    
    "2. Update CODEOWNERS": "If changes exist, copy the new content and update the associated CODEOWNERS file.",
    
    "2b. Edit and replace lines": "Edit the file, replace outdated lines, and create a PR to update it.",
    
    "3. Update Docs": "Run merge-report to find impacted docs. Update metadata to trigger rebuilds.",
    
    "Update maintenance repo": "Update the Maintenance Repo files so the next person has the latest version.",
    
    "Questions?": "If you have any questions, feel free to reach out.",
    
    "Resources": "More info at github.com/sdgilley/content-maintenance and the Microsoft Learn platform manual."
}
