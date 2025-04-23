FROM public.ecr.aws/lambda/python:3.12

# Install OS deps (for lxml, matplotlib, docx)
RUN yum install -y gcc libxml2-devel libxslt-devel atlas-devel freetype-devel libpng-devel && \
    pip install --upgrade pip

# Copy your handler
COPY lambda_function.py ${LAMBDA_TASK_ROOT}/

# Install Python deps
RUN pip install \
    python-docx \
    matplotlib \
    lxml \
    boto3 \
    --target "${LAMBDA_TASK_ROOT}"

# Set the CMD to your handler entrypoint
CMD ["lambda_function.lambda_handler"]
