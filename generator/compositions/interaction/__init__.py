"""Interaction-sensitive composition generators.

Each test case composes operations that interact causally -- one operation's
outcome changes the context for the next. Three interaction types:

- Prerequisite: Operation A creates the object that operation B targets.
- Scope-changing: Operation A changes document structure; operation B must
  navigate to elements that didn't exist before A.
- Layout-coupled: Geometric operation A changes the spatial context that
  geometric operation B depends on.

Overwrite/reset pilot cases are included in the prerequisite generator.
"""
