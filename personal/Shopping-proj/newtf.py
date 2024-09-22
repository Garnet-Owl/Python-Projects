# Tensorflow Introduction

import os
os.environ["TF_CPP_MIN_LOG_LEVEL"] = "2"
import tensorflow as tf

# Basic Initialization of Tensors

x = tf.constant(4, shape=(1, 1), dtype=tf.float32)
x = tf.constant([[1, 2, 3], [4, 5, 6]])

x = tf.ones((3, 3))
x = tf.zeros((3, 3))
x = tf.eye(3)  # I for the identity matrix
x = tf.random.normal((3, 3), mean=0, stddev=1)
x = tf.random.uniform((1, 3), minval=0, maxval=1)
x = tf.range(9)
x = tf.range(start=0, limit=10, delta=2)  # delta is a step

print(x)
