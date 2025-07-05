# SnapAssist

Summon intelligence in a snap

![](docs/intro.gif)

## Getting started

### 1. Clone Repository

```shell
git clone git@github.com:namuan/snap-assist.git
```

### 2. Set Up Your Development Environment

Then, install the environment and the pre-commit hooks with

**Pre-requisites**

- uv
- python 3.9+

```shell
make install
```

This will also generate your `uv.lock` file

### 3. Run the pre-commit hooks

Initially, the CI/CD pipeline might be failing due to formatting issues. To resolve those run:

```shell
make check
```

You are now ready to start development on your project!

### 4. Run the application

```shell
make run
```

### 5. Alfred Integration

![](docs/alfred-workflow.png)

![](docs/universal-action.png)

![](docs/hotekey.png)

![](docs/run-script.png)
