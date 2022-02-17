# ComCastExample
Sample project for StackOverflow question 71035209

## Behavior on a machine with office installed

On a machine with the mshtml interop assemblies in the gac the output is this:

```text
Element has Tag INPUT and value John
Element is NOT a frame element
```

## Behavior on a machine without interop assemblies

Deployed on a machine without mshtml assemblies in the GAC type loading succeeds,
but it now allows an invalid cast:

```text
Element has Tag INPUT and value John
Element is a frame element
```
