import requests
from os.path import expanduser

def get_github_script(user, repo, branch, file):
    """Returns the contents of a file on GitHub."""
    url = 'https://raw.githubusercontent.com/{0}/{1}/{2}/{3}'.format(
        user, repo, branch, file
    )
    return requests.get(url).text

def save_on_desktop(filename, contents):
    path_to_file = expanduser('~') + '\\Desktop\\' + filename
    with open(path_to_file, 'w') as write_file:
        write_file.write(contents)

def main():
    branch = input("Which branch do you want? (Press Enter if you don't know) ")
    if not branch:  # Equivalent to branch == ''
        branch = 'master'
    script = get_github_script('carter-lavering', 'stocks', branch, 'stock_get.py')
    if script == 'Not Found':
        raise FileNotFoundError('Invalid URL to the file. Check for typos in'
            ' the branch.')
    save_on_desktop('stock_get.py', script)
    print('Latest version of script successfully saved to the desktop.')

if __name__ == '__main__':
    main()